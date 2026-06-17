# docx-editor — Realtime Collaboration

A minimal demo of multi-user collaborative editing on top of `@eigenpal/docx-editor-react` using [Yjs](https://yjs.dev), [`y-prosemirror`](https://github.com/yjs/y-prosemirror), and [`y-webrtc`](https://github.com/yjs/y-webrtc).

**No backend required.** Peers find each other through Yjs's public WebRTC signaling servers and sync directly browser-to-browser.

## Try it

```bash
bun install
bun run dev
```

Open <http://localhost:5273>, then click **Share link** and paste the URL into a second browser window. Type in either window — edits, selections, and avatars sync live.

## How it works

Five pieces:

1. **`externalContent` prop** tells the editor to treat its `document` prop as a schema seed only and skip the mount-time content load. `ySyncPlugin` populates ProseMirror from the shared `Y.Doc` instead.
2. **`externalPlugins`** receives `ySyncPlugin`, `yCursorPlugin(awareness)`, and `yUndoPlugin()` so Yjs owns the document state, remote cursors, and history. **Tracked changes sync automatically** through this — `insertion`/`deletion` mark attrs (author, date, revision id) ride along with the synced PM tree. **Remote cursors and selection-range highlights** also surface automatically via the editor's PM-decoration forwarding layer.
3. **Awareness** (Yjs's ephemeral state channel) carries each user's name, color, and selection. The `AvatarStack` in the title bar reads `provider.awareness.getStates()` and renders connected users.
4. **Controlled `comments` prop** + a `Y.Array<Comment>` on the same `Y.Doc`. PM only carries the comment range markers; the thread metadata (text, author, replies, resolved status) lives in the Y.Array, mirrored into React state and pushed back through `onCommentsChange`.
5. **Controlled `styleDefinitions` prop** + a `Y.Map` entry on the same `Y.Doc`. The body PM syncs paragraph content through `y-prosemirror`; the document style package (`word/styles.xml`) syncs separately through `styles.get('definitions')`, mirrored into React state and pushed back through `onStyleDefinitionsChange`. The demo persists the encoded Y.Doc to localStorage, so style edits survive a local reload along with body/comments.

```tsx
<DocxEditor
  document={createEmptyDocument()} // schema seed only
  externalContent // skip the load — Yjs owns content
  externalPlugins={[ySync, yCursor, yUndo]}
  comments={comments} // mirrored from Y.Array<Comment>
  onCommentsChange={setComments} // writes back into Y.Array
  styleDefinitions={styleDefinitions} // mirrored from Y.Map('styles')
  onStyleDefinitionsChange={setStyleDefinitions} // writes back into Y.Map
  author={user.name} // attribution for comments / track changes
  renderTitleBarRight={() => <AvatarStack users={users} />}
/>
```

## Files

| File                      | What it does                                                                                                                                 |
| ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/App.tsx`             | Wires identity, room, collaboration hook, and renders `DocxEditor`                                                                           |
| `src/useCollaboration.ts` | Sets up `Y.Doc`, local encoded-state persistence, `WebrtcProvider`, awareness, the y-prosemirror plugins, comments Y.Array, and styles Y.Map |
| `src/AvatarStack.tsx`     | Overlapping circular avatars in the title bar                                                                                                |
| `src/identity.ts`         | Per-tab user identity (sessionStorage) and room id from URL hash                                                                             |

## Caveats

This is a **demo**, not a production-ready collab template. Known gaps:

- **Comments can lose data on concurrent edits.** The Y.Array sync is naive replace-all (`delete(0, length); push(next)` inside a transact). If two peers add a comment in the same instant, whichever transact lands second wipes the other's additions. For production, use `Y.Map<id, Comment>` keyed by comment id — single-key writes resolve concurrent edits cleanly.
- **Comment IDs can collide between peers.** Comment IDs come from a module-level scalar starting at 1, and the demo seeds with an empty document so the load-time bump never fires. First comment from each peer gets `id: 1`. Tracked by [#257](https://github.com/eigenpal/docx-editor/issues/257).
- **Tracked-change accept/reject races.** Two peers accepting/rejecting the same change at the same instant produce two PM transactions over overlapping ranges. Yjs picks an ordering and the loser's intent is silently dropped — no conflict UI.
- **The public WebRTC signaling servers are best-effort.** For a stable connection, deploy [`y-websocket`](https://github.com/yjs/y-websocket), [PartyKit](https://www.partykit.io/), [Liveblocks](https://liveblocks.io/), or [Hocuspocus](https://tiptap.dev/hocuspocus).
- **Persistence is local-only.** This demo writes the encoded Y.Doc state to localStorage for reload durability in the current browser. Use [`y-indexeddb`](https://github.com/yjs/y-indexeddb) or a server-side persistence layer for shared long-term persistence.
- **Loading an existing `.docx` into a live room is non-trivial.** The source-of-truth swap from `document`/`documentBuffer` to the `Y.Doc` needs to happen exactly once and only on a designated peer. Out of scope for this demo.
