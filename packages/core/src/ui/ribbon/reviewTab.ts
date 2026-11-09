/** Review tab: Threaded Comments, Notes, Protect. */
import * as cmd from '../commands';
import { TINT, type Command, type Ctx, type Tab } from './types';

function threadAtCursor(ctx: Ctx) {
  const { r, c } = ctx.doc.selection.active;
  const m = ctx.doc.sheet.mergeAt(r, c);
  return ctx.doc.sheet.commentAt(m?.r0 ?? r, m?.c0 ?? c);
}

const hasThreads = (ctx: Ctx) => ctx.doc.wb.sheets.some((s) => s.comments.some((t) => t.kind === 'threaded'));

export const R = {
  newComment: { id: 'new-comment', label: 'New Comment', icon: 'add_comment', tint: TINT.purple, shortcut: 'Ctrl+Alt+M', run: (ctx) => cmd.newComment(ctx), disabled: (ctx) => ctx.state.readOnly } as Command,
  deleteComment: {
    id: 'delete-comment',
    label: 'Delete',
    tip: 'Delete Comment',
    icon: 'delete',
    disabled: (ctx) => threadAtCursor(ctx)?.kind !== 'threaded',
    run: (ctx) => {
      const t = threadAtCursor(ctx);
      if (t?.kind === 'threaded') ctx.doc.deleteComment(t.comments[0]?.id ?? t.id);
    },
  } as Command,
  previous: { id: 'previous-comment', label: 'Previous', tip: 'Previous Comment', icon: 'navigate_before', disabled: (ctx) => !hasThreads(ctx), run: (ctx) => cmd.stepComment(ctx, -1) } as Command,
  next: { id: 'next-comment', label: 'Next', tip: 'Next Comment', icon: 'navigate_next', disabled: (ctx) => !hasThreads(ctx), run: (ctx) => cmd.stepComment(ctx, 1) } as Command,
  showComments: { id: 'show-comments', label: 'Show Comments', icon: 'forum', tint: TINT.purple, checked: (ctx) => ctx.panel === 'comments', run: (ctx) => ctx.setPanel(ctx.panel === 'comments' ? null : 'comments') } as Command,
  notes: {
    id: 'notes',
    label: 'Notes',
    icon: 'sticky_note_2',
    tint: TINT.orange,
    menu: (ctx) => {
      const t = threadAtCursor(ctx);
      const note = t?.kind === 'note';
      return [
        { title: note ? 'Edit Note' : 'New Note', icon: 'edit_note', shortcut: 'Shift+F2', testId: 'note-edit', disabled: ctx.state.readOnly || t?.kind === 'threaded', onClick: () => cmd.editNote(ctx) },
        { title: 'Delete Note', icon: 'delete', disabled: !note, onClick: () => cmd.deleteNote(ctx) },
        { type: 'line' },
        { title: 'Previous Note', icon: 'navigate_before', onClick: () => stepNote(ctx, -1) },
        { title: 'Next Note', icon: 'navigate_next', onClick: () => stepNote(ctx, 1) },
        { type: 'line' },
        { title: 'Convert to Comments', icon: 'forum', disabled: !ctx.doc.sheet.comments.some((x) => x.kind === 'note'), onClick: () => convertNotes(ctx) },
      ];
    },
  } as Command,
  protectSheet: {
    id: 'protect-sheet',
    label: 'Protect Sheet',
    icon: 'lock',
    tint: TINT.orange,
    checked: (ctx) => ctx.state.sheetProtected,
    run: (ctx) => ctx.openDialog({ kind: ctx.state.sheetProtected ? 'unprotect-sheet' : 'protect-sheet' }),
    disabled: (ctx) => ctx.state.readOnly,
  } as Command,
  protectWorkbook: {
    id: 'protect-workbook',
    label: 'Protect Workbook',
    icon: 'key',
    tint: TINT.orange,
    checked: (ctx) => ctx.state.workbookProtected,
    run: (ctx) => {
      if (!ctx.state.workbookProtected) {
        ctx.openDialog({ kind: 'protect-workbook' });
        return;
      }
      void (async () => {
        if (ctx.doc.unprotectWorkbook('')) return;
        const pw = await ctx.prompt('Password:', '', { title: 'Unprotect Workbook', ok: 'OK' });
        if (pw !== null && !ctx.doc.unprotectWorkbook(pw)) ctx.notify('The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.');
      })();
    },
    disabled: (ctx) => ctx.state.readOnly,
  } as Command,
  lockCell: {
    id: 'lock-cell',
    label: 'Lock Cell',
    icon: 'lock_open',
    checked: (ctx) => ctx.doc.styleAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c).locked !== false,
    run: (ctx) => cmd.format(ctx, { locked: ctx.doc.styleAt(ctx.doc.selection.active.r, ctx.doc.selection.active.c).locked === false }, 'Lock cell'),
    disabled: (ctx) => ctx.state.readOnly || ctx.state.sheetProtected,
  } as Command,
};

function stepNote(ctx: Ctx, dir: 1 | -1): void {
  const notes = [...ctx.doc.sheet.comments].filter((t) => t.kind === 'note').sort((a, b) => a.r - b.r || a.c - b.c);
  if (!notes.length) return;
  const { r, c } = ctx.doc.selection.active;
  const key = (x: { r: number; c: number }) => x.r * 1e5 + x.c;
  const k = key({ r, c });
  const next = dir > 0 ? notes.find((t) => key(t) > k) ?? notes[0] : [...notes].reverse().find((t) => key(t) < k) ?? notes[notes.length - 1];
  ctx.doc.selectCell(next.r, next.c);
  ctx.grid?.scrollIntoView(next.r, next.c);
  ctx.openNote(next.r, next.c);
}

/** Notes become threaded comments with the note text as the first comment. */
function convertNotes(ctx: Ctx): void {
  const { doc } = ctx;
  const notes = doc.sheet.comments.filter((t) => t.kind === 'note');
  doc.transact('Convert to Comments', () => {
    for (const t of notes) {
      const text = t.comments.map((x) => x.text).join('\n').replace(/^[^:\n]{1,40}:\n/, '');
      doc.setNote(t.r, t.c, '');
      doc.addComment(t.r, t.c, text, t.comments[0]?.author || doc.author);
    }
  });
}

export const reviewTab: Tab = {
  id: 'review',
  label: 'Review',
  groups: () => [
    {
      id: 'comments',
      label: 'Comments',
      icon: 'forum',
      controls: [
        { type: 'large', cmd: R.newComment },
        { type: 'stack', items: [{ type: 'small', cmd: R.deleteComment }, { type: 'small', cmd: R.previous }, { type: 'small', cmd: R.next }] },
        { type: 'large', cmd: R.showComments },
      ],
    },
    { id: 'notes', label: 'Notes', icon: 'sticky_note_2', controls: [{ type: 'large', cmd: R.notes }] },
    {
      id: 'protect',
      label: 'Protect',
      icon: 'lock',
      controls: [{ type: 'large', cmd: R.protectSheet }, { type: 'large', cmd: R.protectWorkbook }, { type: 'large', cmd: R.lockCell }],
    },
  ],
  toolbar: () => [
    { type: 'small', cmd: R.newComment },
    { type: 'small', cmd: R.deleteComment },
    { type: 'small', cmd: R.previous },
    { type: 'small', cmd: R.next },
    { type: 'small', cmd: R.showComments },
    { type: 'small', cmd: R.notes },
    { type: 'small', cmd: R.protectSheet },
    { type: 'small', cmd: R.protectWorkbook },
  ],
};
