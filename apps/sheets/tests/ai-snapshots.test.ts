import { describe, expect, it, vi } from 'vitest'
import type { SheetSnapshot } from '../src/renderer/ai/AiChatPanel'

describe('AI Snapshot and Rollback Lifecycle', () => {
  it('deduplicates snapshots during a multi-step agent run so only one snapshot is created per prompt', () => {
    const snapshots: SheetSnapshot[] = []
    let aiRunSnapshotCreated = false
    const instruction =
      'Create a complete financial balance sheet table with assets, liabilities, equity'

    function recordSnapshot(_stepName: string) {
      if (!aiRunSnapshotCreated) {
        aiRunSnapshotCreated = true
        snapshots.unshift({
          id: `snap-${Date.now()}-${Math.random()}`,
          label: instruction.slice(0, 40),
          time: '06:50 PM',
        })
      }
    }

    // Step 1: create table
    recordSnapshot('create table')
    expect(snapshots).toHaveLength(1)
    expect(snapshots[0]?.label).toBe(instruction.slice(0, 40))

    // Step 2: insert formulas (same agent run)
    recordSnapshot('insert formulas')
    expect(snapshots).toHaveLength(1)

    // Step 3: format header cells (same agent run)
    recordSnapshot('format cells')
    expect(snapshots).toHaveLength(1)

    // Step 4: adjust widths (same agent run)
    recordSnapshot('adjust widths')
    expect(snapshots).toHaveLength(1)

    // Now a NEW user prompt arrives
    aiRunSnapshotCreated = false
    const nextInstruction = 'Make the header font bold'
    function recordNextSnapshot() {
      if (!aiRunSnapshotCreated) {
        aiRunSnapshotCreated = true
        snapshots.unshift({
          id: `snap-${Date.now()}-${Math.random()}`,
          label: nextInstruction.slice(0, 40),
          time: '06:55 PM',
        })
      }
    }

    recordNextSnapshot()
    expect(snapshots).toHaveLength(2)
    expect(snapshots[0]?.label).toBe('Make the header font bold')
    expect(snapshots[1]?.label).toBe(instruction.slice(0, 40))
  })

  it('unwinds rollback stack properly when rolling back to an older snapshot', () => {
    let snapshots: SheetSnapshot[] = [
      { id: 'snap-3', label: 'Prompt 3', time: '07:00 PM' },
      { id: 'snap-2', label: 'Prompt 2', time: '06:55 PM' },
      { id: 'snap-1', label: 'Prompt 1', time: '06:50 PM' },
    ]

    const undoMock = vi.fn()

    function rollback(snapshot: SheetSnapshot) {
      const index = snapshots.findIndex((s) => s.id === snapshot.id)
      if (index === -1) return
      const undosToPerform = index + 1
      for (let i = 0; i < undosToPerform; i++) {
        undoMock()
      }
      snapshots = snapshots.slice(index + 1)
    }

    // Rollback to Prompt 2 (index 1 in stack: should undo Prompt 3 and Prompt 2)
    rollback(snapshots[1]!)
    expect(undoMock).toHaveBeenCalledTimes(2)
    expect(snapshots).toHaveLength(1)
    expect(snapshots[0]?.id).toBe('snap-1')

    // Now rollback to Prompt 1 (index 0 in remaining stack)
    undoMock.mockClear()
    rollback(snapshots[0]!)
    expect(undoMock).toHaveBeenCalledTimes(1)
    expect(snapshots).toHaveLength(0)
  })

  it('defensively filters out duplicate snapshot IDs in display logic', () => {
    const rawSnapshots: SheetSnapshot[] = [
      { id: 'snap-1', label: 'Create balance sheet', time: '06:50 PM' },
      { id: 'snap-1', label: 'Create balance sheet', time: '06:50 PM' }, // duplicate ID
      { id: 'snap-2', label: 'Add chart', time: '06:55 PM' },
    ]

    const seenIds = new Set<string>()
    const displaySnapshots: SheetSnapshot[] = []
    for (const s of rawSnapshots) {
      if (seenIds.has(s.id)) continue
      seenIds.add(s.id)
      displaySnapshots.push(s)
    }

    expect(displaySnapshots).toHaveLength(2)
    expect(displaySnapshots[0]?.id).toBe('snap-1')
    expect(displaySnapshots[1]?.id).toBe('snap-2')
  })
})
