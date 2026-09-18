import { utils, type WorkSheet } from 'xlsx';

export interface ResultCell {
  text: string;
  rowspan: number;
  colspan: number;
}

// Keep covered cells out of the HTML grid and render only each merge's anchor.
export function resultRows(sheet: WorkSheet): ResultCell[][] {
  if (!sheet['!ref']) return [];
  const range = utils.decode_range(sheet['!ref']);
  const merges = sheet['!merges'] ?? [];
  for (const merge of merges) {
    range.s.r = Math.min(range.s.r, merge.s.r);
    range.s.c = Math.min(range.s.c, merge.s.c);
    range.e.r = Math.max(range.e.r, merge.e.r);
    range.e.c = Math.max(range.e.c, merge.e.c);
  }
  const rows: ResultCell[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cells: ResultCell[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const merge = merges.find(({ s, e }) =>
        r >= s.r && r <= e.r && c >= s.c && c <= e.c);
      if (merge && (r !== merge.s.r || c !== merge.s.c)) continue;
      const cell = sheet[utils.encode_cell({ r, c })];
      cells.push({
        text: cell ? utils.format_cell(cell) : '',
        rowspan: merge ? merge.e.r - merge.s.r + 1 : 1,
        colspan: merge ? merge.e.c - merge.s.c + 1 : 1,
      });
    }
    rows.push(cells);
  }
  return rows;
}
