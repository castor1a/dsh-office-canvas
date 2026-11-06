import { fileURLToPath } from 'node:url';

export const inject = ['connection', 'agents'];
const defaults = {
  pythonPath: 'C:\\Users\\ftt20\\.dsh\\dsh-runtimes\\dsh-next-primary-runtime\\dependencies\\python\\python.exe',
  nodePath: 'H:\\DSH\\DSH NEXT\\resources\\runtime\\x64\\primary-runtime\\dependencies\\node\\bin\\node.exe',
  libreofficeModulePath: 'H:\\DSH\\DSH NEXT\\resources\\app.asar.unpacked\\node_modules\\@deepseek-ai\\libreoffice-kit\\lib\\index.js'
};
const backend = fileURLToPath(new URL('./backend.py', import.meta.url));
const renderer = fileURLToPath(new URL('./render.mjs', import.meta.url));
const quote = s => "'" + s.replaceAll("'", "''") + "'";

export function apply(ctx, config = {}) {
  config = {...defaults, ...config};
  async function run(request, agent, signal) {
    if (!agent || !request || !['read', 'save', 'preview'].includes(request.action) || typeof request.path !== 'string' || !request.path.trim() || request.path.includes('\0')) throw Error('Invalid Office request');
    const ac = agent.ctx;
    const policyService = ac.get('sandboxPolicy');
    const shell = ac.get('shell');
    if (!policyService || !shell) throw Error('This conversation does not provide a sandboxed shell');
    const policy = policyService.resolve({session: agent.session});
    if (request.action === 'save' && policy.mode === 'read-only') throw Error('File access denied under read-only mode');
    if (request.pages !== undefined && (!Array.isArray(request.pages) || !request.pages.length || request.pages.length > 5 || request.pages.some(p => !Number.isSafeInteger(p) || p < 1))) throw Error('Preview supports one to five positive page numbers');
    const preview = request.action === 'preview';
    const payload = {action: request.action, path: request.path, workspaceRoot: policy.workspaceRoot};
    for (const key of ['expectedHash', 'expectedOutputHash', 'operations', 'outputPath', 'pages']) if (request[key] !== undefined) payload[key] = request[key];
    if (preview) payload.libreofficeModulePath = config.libreofficeModulePath;
    const stdin = JSON.stringify(payload);
    if (Buffer.byteLength(stdin, 'utf8') > 1024 * 1024) throw Error('Office request exceeds 1 MiB');
    const command = `& ${quote(preview ? config.nodePath : config.pythonPath)} ${quote(preview ? renderer : backend)}; exit $LASTEXITCODE`;
    const spec = shell.resolve({command, stdin, workdir: policy.workspaceRoot, sandboxPolicy: policy, signal, timeoutMs: 120000, stdoutMaxBytes: preview ? 32 * 1024 * 1024 : 16 * 1024 * 1024});
    const result = await (await shell.execute(spec)).result();
    if (result.sandbox?.denied) throw Error(`File access denied under ${policy.mode} mode`);
    if (result.aborted) throw (signal?.reason ?? Error('Cancelled'));
    if (result.timedOut) throw Error('Office operation timed out');
    if (result.exitCode !== 0) throw Error(result.stderr.text || `Office backend exited ${result.exitCode}`);
    if (result.stdout.truncated) throw Error('Office response exceeded its size limit');
    return JSON.parse(result.stdout.text.replace(/^\uFEFF/, ''));
  }
  ctx.connection.fetch.register({
    path: '/api/office-canvas.operation', methods: ['POST'], requestBody: 'buffered',
    async fetch(request) {
      try {
        const body = await request.json();
        if (typeof body.sessionId !== 'string') throw Error('A session is required');
        const agent = ctx.agents.get(body.sessionId);
        if (!agent) throw Error('Open this file in an active conversation first');
        const value = await run(body.request, agent, request.signal);
        return Response.json(value);
      } catch (error) {
        return Response.json({ok: false, error: {code: 'office-operation-failed', message: error instanceof Error ? error.message : String(error)}}, {status: 400});
      }
    }
  });
  const registrations = new Map();
  const mountAgent = agent => {
    if (registrations.has(agent.id)) return;
    const ac = agent.ctx;
    const dispose = ac.inject(['commands', 'tools', 'shell', 'sandboxPolicy'], scope => {
  scope.commands.register({name: 'office-canvas', description: 'Read, preview and edit DOCX/PPTX files', input: {hint: 'Office operation JSON'}, recordInput: false,
    async handler({agent, rawInput, signal}) {
      try { return {kind: 'success', text: JSON.stringify(await run(JSON.parse(rawInput), agent, signal))}; }
      catch (error) { return {kind: 'error', text: error instanceof Error ? error.message : String(error)}; }
    }
  });
  scope.tools.register({name: 'office_canvas_edit', description: 'Read/edit DOCX and PPTX while preserving untouched OOXML parts. Read first for hash and stable IDs. Save uses operations text/format/move, expectedHash, optional outputPath. Preview renders up to five pages. Paths must be in this session workspace.',
    parameters: {type: 'object', properties: {request: {type: 'string', description: 'JSON {action:read|save|preview,path,expectedHash?,operations?:[{type:text|format|move,id,...}],outputPath?,expectedOutputHash?,pages?:[1]}'}}, required: ['request'], additionalProperties: false},
    output: {schema: {type: 'string'}, render: (_args, value) => [{type: 'text', text: value}]}, isConcurrencySafe: () => false,
    async execute(args, exec) { return JSON.stringify(await run(JSON.parse(args.request), exec.agent, exec.signal)); }
  });
    });
    registrations.set(agent.id, () => dispose.dispose());
  };
  ctx.on('agent/created', ({agent}) => mountAgent(agent));
  ctx.on('agent/disposed', ({agent}) => { registrations.get(agent.id)?.(); registrations.delete(agent.id); });
  for (const agent of ctx.agents.list()) mountAgent(agent);
  ctx.effect(() => () => { for (const dispose of registrations.values()) dispose(); registrations.clear(); });
}
