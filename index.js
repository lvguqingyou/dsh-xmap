/**
 * dsh-xmap — DeepSeek Harness tool plugin wrapping the xmap PowerShell engine.
 *
 * xmap (see ./xmap/README.md) is a generic Excel schema-map tool: the model
 * writes a small column-mapping "spec" (JSON) and the engine reads the whole
 * sheet out of context via Excel COM. This plugin registers one tool per xmap
 * subcommand and executes the PowerShell engine through the `ctx.shell`
 * capability seam (the same executor that backs the built-in pwsh tool), so
 * sandboxing and the per-session file policy apply to every call.
 *
 * Design follows the official tool cookbook (`docs/cookbook/adding-a-tool`)
 * and mirrors the foreground path of `packages/shell/tool-pwsh`.
 */

import { fileURLToPath } from 'node:url'
import { isAbsolute, resolve as resolvePath } from 'node:path'
import z from '@deepseek-ai/schemastery'
import { defineTool } from '@deepseek-ai/dsh-tools'
import { ESCALATION_TARGETS, approveEscalation, validateEscalationArgs } from '@deepseek-ai/dsh-sandbox'

export const name = 'dsh-xmap'
export const inject = ['tools', 'shell']

/**
 * Plugin configuration. No hardcoded tunables: values a deployment may want to
 * change are declared here and editable from cordis.yml `config:`.
 * - specDir: spec-library directory; empty defaults to the specs bundled with this package.
 * - timeoutMs: foreground timeout in milliseconds for one Excel COM run.
 *
 * Runtime configuration schema (Schemastery).
 */
export const Config = z.object({
  specDir: z.string().default(''),
  timeoutMs: z.number().default(300000),
})

const PACKAGED_SCRIPT = fileURLToPath(new URL('./xmap/xmap.ps1', import.meta.url))
const PACKAGED_SPECS = fileURLToPath(new URL('./xmap/specs', import.meta.url))

/** Single-quote a PowerShell literal (doubling embedded quotes). */
function psQuote(value) {
  return `'${String(value).replace(/'/g, "''")}'`
}

/** Non-empty check for values the schema DSL cannot express. */
function requireNonEmpty(args, key) {
  const v = args[key]
  if (v === undefined || String(v).trim() === '') {
    throw new Error(`invalid ${key}: expected a non-empty value`)
  }
  return String(v).trim()
}

/** The canonical foreground result shape returned by every xmap tool. */
function canonicalResult(result) {
  return {
    exitCode: result.exitCode ?? -1,
    timedOut: result.timedOut,
    aborted: result.aborted,
    stdout: result.stdout.text,
    stderr: result.stderr.text,
  }
}

/** Model-visible rendering of one canonical result. */
function renderResult(_args, value) {
  const parts = []
  if (value.stdout) parts.push(value.stdout.replace(/\n+$/, ''))
  if (value.stderr) parts.push(`[stderr]\n${value.stderr.replace(/\n+$/, '')}`)
  if (value.exitCode !== 0) parts.push(`[exit code: ${value.exitCode}]`)
  const text = parts.join('\n')
  return [{ type: 'text', text: text === '' ? '(no output)' : text }]
}

export function apply(ctx, config = {}) {
  const scriptPath = PACKAGED_SCRIPT
  const specDir = config.specDir && String(config.specDir).trim() ? String(config.specDir).trim() : PACKAGED_SPECS
  const defaultTimeoutMs = config.timeoutMs ?? 300000

  // The web composition mounts a confining shell executor, so the standing
  // per-session sandbox policy must be resolved for each call (mirrors
  // tool-pwsh/tool-bash). Escalation params are advertised only when an
  // escalation target exists.
  const defaultMode = ctx.shell.sandboxMode
  const escalationModes = defaultMode === undefined ? [] : ESCALATION_TARGETS
  const sandboxPolicy = defaultMode === undefined ? undefined : ctx.get('sandboxPolicy')
  if (defaultMode !== undefined && sandboxPolicy === undefined) {
    throw new Error('dsh-xmap: the mounted shell executor confines but ctx.sandboxPolicy is missing')
  }

  const resolvePolicy = (exec) =>
    sandboxPolicy?.resolve(exec.agent === undefined ? {} : { session: exec.agent.session })

  const approveMode = async (mode, justification, exec, toolName, standingPolicy) => {
    if (escalationModes.length === 0) {
      throw new Error('sandbox_permissions is not available in this composition (no sandboxing executor to escalate)')
    }
    const effectiveMode = standingPolicy.mode
    return approveEscalation(
      { requestedMode: mode, justification, effectiveMode, subject: 'command' },
      {
        approver: ctx.get('approval'),
        agent: exec.agent,
        callId: exec.callId,
        toolName,
        signal: exec.signal,
      },
    )
  }

  const resolveWorkdir = (modelWorkdir, exec) => {
    const headerCwd = exec.agent?.session.header.cwd
    if (modelWorkdir === undefined) return headerCwd
    if (headerCwd !== undefined && !isAbsolute(modelWorkdir)) return resolvePath(headerCwd, modelWorkdir)
    return modelWorkdir
  }

  /** Run one xmap subcommand through the shell seam and return its canonical result. */
  const runXmap = async (exec, toolName, commandTokens, args) => {
    const standingPolicy = resolvePolicy(exec)
    const approvedMode =
      args.sandbox_permissions !== undefined && args.justification !== undefined
        ? await approveMode(args.sandbox_permissions, args.justification, exec, toolName, standingPolicy)
        : undefined
    const policy =
      approvedMode === undefined
        ? standingPolicy
        : { ...standingPolicy, mode: approvedMode }
    const workdir = resolveWorkdir(args.workdir, exec)
    const command = `& ${psQuote(scriptPath)} ${commandTokens.join(' ')}`
    const request = {
      command,
      ...workdir !== undefined ? { workdir } : {},
      timeoutMs: args.timeoutMs ?? defaultTimeoutMs,
      ...policy !== undefined ? { sandboxPolicy: policy } : {},
    }
    const result = await ctx.shell.run(ctx.shell.resolve({ ...request, signal: exec.signal }))
    if (result.aborted) {
      const error = new Error('xmap tool call aborted')
      error.name = 'AbortError'
      throw error
    }
    return canonicalResult(result)
  }

  /** Common parameter surface for escalation + workdir on every tool. */
  const sharedParams = (escalationOn) => ({
    workdir: { type: 'string', description: 'Working directory for this call. Defaults to the session workspace; a relative path is resolved against it.' },
    timeoutMs: { type: 'number', description: `Timeout in milliseconds (default ${defaultTimeoutMs}); the executor caps it.` },
    ...escalationOn ? {
      sandbox_permissions: {
        type: 'string',
        enum: [...escalationModes],
        description: 'The wider sandbox mode this command needs. Only valid as a one-shot retry of a command the sandbox just denied; requires justification and user approval.',
      },
      justification: {
        type: 'string',
        description: 'Required with sandbox_permissions: one sentence for the user explaining why this exact command needs the wider access.',
      },
    } : {},
  })

  const register = (tool) => {
    ctx.tools.register(defineTool({
      name: tool.name,
      description: tool.description,
      parameters: { ...tool.parameters, ...sharedParams(escalationModes.length > 0) },
      output: {
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            exitCode: { type: 'integer', required: true },
            timedOut: { type: 'boolean', required: true },
            aborted: { type: 'boolean', required: true },
            stdout: { type: 'string', required: true },
            stderr: { type: 'string', required: true },
          },
        },
        render: renderResult,
      },
      async execute(args, exec) {
        const tokens = tool.build(args)
        return runXmap(exec, tool.name, tokens, args)
      },
    }))
  }

  // ------------------------------------------------------------------
  // sheets — list every sheet in a workbook (name + rows + cols)
  // ------------------------------------------------------------------
  register({
    name: 'xmap_sheets',
    description: 'List all worksheets of an .xlsx workbook with their used row/column counts, via Excel COM. '
      + 'Use before extracting to find the exact sheet name. '
      + 'Example: xmap_sheets file="C:\\data\\book.xlsx".',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook (Windows form, e.g. C:\\...).' },
    },
    build(args) {
      return ['sheets', psQuote(requireNonEmpty(args, 'file'))]
    },
  })

  // ------------------------------------------------------------------
  // headers — preview header rows / sample rows of one sheet
  // ------------------------------------------------------------------
  register({
    name: 'xmap_headers',
    description: 'Print the first N rows (default 5) of one worksheet as text so the structure/headers can be inspected '
      + 'before writing an extraction spec. Example: xmap_headers file="C:\\b.xlsx" sheet="Sheet1" rows=5.',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook.' },
      sheet: { type: 'string', required: true, description: 'Worksheet name exactly as listed by xmap_sheets.' },
      rows: { type: 'number', description: 'How many rows to preview (default 5).' },
    },
    build(args) {
      const tokens = ['headers', psQuote(requireNonEmpty(args, 'file')), '-Sheet', psQuote(requireNonEmpty(args, 'sheet'))]
      if (args.rows !== undefined) {
        if (!Number.isInteger(args.rows) || args.rows < 1) throw new Error('invalid rows: expected a positive integer')
        tokens.push('-Rows', String(args.rows))
      }
      return tokens
    },
  })

  // ------------------------------------------------------------------
  // hash — header fingerprint used for spec reuse
  // ------------------------------------------------------------------
  register({
    name: 'xmap_hash',
    description: 'Compute the header fingerprint (MD5 of the first rows) of a worksheet. Equal fingerprints mean the same '
      + 'column layout, so an existing spec in the spec library will be reused automatically by xmap_extract. '
      + 'Example: xmap_hash file="C:\\b.xlsx" sheet="Sheet1".',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook.' },
      sheet: { type: 'string', required: true, description: 'Worksheet name.' },
    },
    build(args) {
      return ['hash', psQuote(requireNonEmpty(args, 'file')), '-Sheet', psQuote(requireNonEmpty(args, 'sheet'))]
    },
  })

  // ------------------------------------------------------------------
  // extract — read the whole sheet into normalized JSONL per a spec
  // ------------------------------------------------------------------
  register({
    name: 'xmap_extract',
    description: 'Extract one worksheet into normalized JSONL records using a column-mapping spec. '
      + 'If `spec` is omitted, xmap auto-resolves a spec from the spec library by the sheet\'s header fingerprint; '
      + 'when none matches, the call fails with the fingerprint so a spec can be written. '
      + 'Example: xmap_extract file="C:\\b.xlsx" sheet="Sheet1" out="C:\\records.jsonl". '
      + 'See the spec JSON shape in the xmap README (headerRows/dataStart/cols/numeric/skipIfNameEmpty).',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook.' },
      sheet: { type: 'string', required: true, description: 'Worksheet name.' },
      spec: { type: 'string', description: 'Optional absolute path to a spec JSON. Omit to auto-match by fingerprint from the spec library (config `specDir`, default the bundled specs folder).' },
      out: { type: 'string', required: true, description: 'Absolute path of the JSONL output file to write.' },
    },
    build(args) {
      const tokens = ['extract', psQuote(requireNonEmpty(args, 'file')), '-Sheet', psQuote(requireNonEmpty(args, 'sheet'))]
      if (args.spec !== undefined) tokens.push('-Spec', psQuote(requireNonEmpty(args, 'spec')))
      tokens.push('-SpecDir', psQuote(specDir), '-Out', psQuote(requireNonEmpty(args, 'out')))
      return tokens
    },
  })

  // ------------------------------------------------------------------
  // verify — counts + numeric sums against the extracted JSONL
  // ------------------------------------------------------------------
  register({
    name: 'xmap_verify',
    description: 'Verify an extraction: prints the record count and the sum of every numeric field declared in the spec, '
      + 'and optionally compares against an in-sheet total row (spec.checkTotal). '
      + 'Example: xmap_verify file="C:\\b.xlsx" sheet="Sheet1" spec="C:\\spec.json" records="C:\\records.jsonl".',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook (needed for checkTotal).' },
      sheet: { type: 'string', required: true, description: 'Worksheet name.' },
      spec: { type: 'string', required: true, description: 'Absolute path to the spec JSON used by the extraction.' },
      records: { type: 'string', required: true, description: 'Absolute path to the JSONL produced by xmap_extract.' },
      report: { type: 'string', description: 'Optional absolute path to write the verification report text.' },
    },
    build(args) {
      const tokens = ['verify', psQuote(requireNonEmpty(args, 'file')), '-Sheet', psQuote(requireNonEmpty(args, 'sheet')),
        '-Spec', psQuote(requireNonEmpty(args, 'spec')), '-Records', psQuote(requireNonEmpty(args, 'records'))]
      if (args.report !== undefined) tokens.push('-Report', psQuote(requireNonEmpty(args, 'report')))
      return tokens
    },
  })

  // ------------------------------------------------------------------
  // save-spec — store a spec into the library under the sheet fingerprint
  // ------------------------------------------------------------------
  register({
    name: 'xmap_save_spec',
    description: 'Store a validated spec JSON into the spec library under this worksheet\'s header fingerprint so later '
      + 'xmap_extract calls on same-layout sheets reuse it with zero rework. '
      + 'Example: xmap_save_spec file="C:\\b.xlsx" sheet="Sheet1" spec="C:\\spec.json".',
    parameters: {
      file: { type: 'string', required: true, description: 'Absolute path to the .xlsx workbook.' },
      sheet: { type: 'string', required: true, description: 'Worksheet name.' },
      spec: { type: 'string', required: true, description: 'Absolute path to the spec JSON to save.' },
    },
    build(args) {
      return ['save-spec', psQuote(requireNonEmpty(args, 'file')), '-Sheet', psQuote(requireNonEmpty(args, 'sheet')),
        '-Spec', psQuote(requireNonEmpty(args, 'spec')), '-SpecDir', psQuote(specDir)]
    },
  })

  // ------------------------------------------------------------------
  // aggregate — group/sum across several extracted JSONL files
  // ------------------------------------------------------------------
  register({
    name: 'xmap_aggregate',
    description: 'Aggregate one or more extracted JSONL files: group by the given fields and sum the given numeric fields. '
      + '`records` takes comma-separated JSONL paths; `groupBy` and `sum` take comma-separated field names that exist in the records. '
      + 'Example: xmap_aggregate records="a.jsonl,b.jsonl" groupBy="name,category" sum="bal_1y,qty_1y" out="agg.jsonl".',
    parameters: {
      records: { type: 'string', required: true, description: 'Comma-separated absolute paths of JSONL files to aggregate.' },
      groupBy: { type: 'string', required: true, description: 'Comma-separated field names to group by.' },
      sum: { type: 'string', required: true, description: 'Comma-separated numeric field names to sum.' },
      out: { type: 'string', description: 'Optional absolute path of the aggregated JSONL to write (otherwise printed).' },
    },
    build(args) {
      const tokens = ['aggregate', '-Records', psQuote(requireNonEmpty(args, 'records')),
        '-GroupBy', psQuote(requireNonEmpty(args, 'groupBy')), '-Sum', psQuote(requireNonEmpty(args, 'sum'))]
      if (args.out !== undefined) tokens.push('-Out', psQuote(requireNonEmpty(args, 'out')))
      return tokens
    },
  })
}
