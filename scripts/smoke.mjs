// Offline smoke test for dsh-xmap: imports the plugin, drives apply() with a
// stub ctx, and exercises every tool's execute() against a stub shell.run().
// Run: node scripts/smoke.mjs   (resolve module from this repo root)

import { apply, name, inject } from '../index.js'

const registered = []
const calls = []

const ctx = {
  shell: {
    // No confining executor in this stub => no escalation surface.
    sandboxMode: undefined,
    resolve: (request) => request,
    run: async (spec) => {
      calls.push(spec.command)
      return {
        exitCode: 0,
        signal: null,
        timedOut: false,
        aborted: false,
        timeoutMs: 1000,
        stdout: { text: 'stub-ok', truncated: false },
        stderr: { text: '', truncated: false },
      }
    },
  },
  get: () => undefined,
  tools: { register: (tool) => registered.push(tool) },
}

console.log('plugin name =', name, '| inject =', inject.join(','))

apply(ctx, { specDir: 'C:\\specs', timeoutMs: 20000 })

console.log('registered tools =', registered.length)
if (registered.length !== 7) throw new Error(`expected 7 tools, got ${registered.length}`)

const names = registered.map((t) => t.name)
console.log('tool names =', names.join(', '))
if (!names.includes('xmap_sheets') || !names.includes('xmap_aggregate')) throw new Error('missing tools')

const exec = { agent: undefined, signal: new AbortController().signal, callId: 'smoke-1' }

for (const tool of registered) {
  const args = argsFor(tool.name)
  const value = await tool.execute(args, exec)
  if (typeof value !== 'object' || value.stdout !== 'stub-ok') throw new Error(`bad result for ${tool.name}`)
  console.log(`  ${tool.name}: execute OK -> exitCode=${value.exitCode}`)
}

function argsFor(toolName) {
  switch (toolName) {
    case 'xmap_sheets': return { file: 'C:\\a.xlsx' }
    case 'xmap_headers': return { file: 'C:\\a.xlsx', sheet: 'S1', rows: 3 }
    case 'xmap_hash': return { file: 'C:\\a.xlsx', sheet: 'S1' }
    case 'xmap_extract': return { file: 'C:\\a.xlsx', sheet: 'S1', out: 'C:\\r.jsonl' }
    case 'xmap_verify': return { file: 'C:\\a.xlsx', sheet: 'S1', spec: 'C:\\s.json', records: 'C:\\r.jsonl' }
    case 'xmap_save_spec': return { file: 'C:\\a.xlsx', sheet: 'S1', spec: 'C:\\s.json' }
    case 'xmap_aggregate': return { records: 'C:\\a.jsonl,C:\\b.jsonl', groupBy: 'name', sum: 'amount' }
    default: throw new Error(`no fixture for ${toolName}`)
  }
}

// Assert one constructed command looks right (xmap_extract with default specDir override).
const extractCmd = calls.find((c) => c.includes("extract 'C:\\a.xlsx'"))
console.log('\nsample extract command:')
console.log(extractCmd)
if (!extractCmd.includes('-SpecDir') || !extractCmd.includes("'C:\\specs'")) {
  throw new Error('extract command missing expected specDir flag')
}
console.log('\nSMOKE OK')
