import { Component, OnInit, NgZone, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormsModule, FormGroup, FormArray, Validators } from '@angular/forms';
import { Workbook, Alignment, Borders } from 'exceljs';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';

@Component({
  standalone: true,
  selector: 'app-test-plan',
  imports: [ CommonModule, ReactiveFormsModule, FormsModule ],
  templateUrl: './test-plan.html',
  styleUrls: ['./test-plan.scss']
})
export class TestPlan implements OnInit {
  testPlanForm!: FormGroup;
  workbook?: XLSX.WorkBook;
  sheetNames: string[] = [];
  selectedSheetName: string | null = null;
  startRow = 19;
  endRow: number | null = null;
  isSidebarOpen = false;
  activeSheetIndex = 0;


  constructor(
    private readonly fb: FormBuilder,
    private readonly ngZone: NgZone,
    private readonly cdr: ChangeDetectorRef
  ) {}
  
  ngOnInit(): void {
    this.testPlanForm = this.fb.group({
      title: ['', Validators.required],
      sheets: this.fb.array([ this.createSheet('Tesztelési forgatókönyv') ])
    });
  }

  createSheet(name = 'Lap 1'): FormGroup {
    return this.fb.group({
      sheetName: [name, Validators.required],
      steps: this.fb.array([ this.createStep() ])
    });
  }

  createStep(): FormGroup {
    return this.fb.group({
      isSubtest:    [false],
      subtestName:  [''],
      subtestColor: ['#929292'],
      function:     [''],
      precondition: [''],
      testData:     [''],
      stepsText:    [''],
      expected:     [''],
      comment:      ['']
    });
  }

  get sheets(): FormArray { return this.testPlanForm.get('sheets') as FormArray; }
  get activeSheetGroup(): FormGroup { return this.sheets.at(this.activeSheetIndex) as FormGroup; }
  get steps(): FormArray { return this.activeSheetGroup.get('steps') as FormArray; }


  addSheet(name?: string): void {
    this.sheets.push(this.createSheet(name ?? `Lap ${this.sheets.length + 1}`));
    this.activeSheetIndex = this.sheets.length - 1;
  }
  removeSheet(idx: number): void {
    if (this.sheets.length <= 1) return;
    this.sheets.removeAt(idx);
    if (this.activeSheetIndex >= this.sheets.length) this.activeSheetIndex = this.sheets.length - 1;
  }
  switchSheet(idx: number): void {
    this.activeSheetIndex = idx;
    this.selection.clear();
  }

  addStep(afterIndex?: number): void {
    const insertAt = ((afterIndex ?? (this.steps.length - 1)) + 1);
    this.steps.insert(insertAt, this.createStep());
  }

  removeStep(index: number): void {
    if (this.selection.size) {
      [...this.selection]
        .sort((a, b) => b - a)
        .forEach(i => {
          if (this.steps.length > 1) {
            this.steps.removeAt(i);
          }
        });
      this.selection.clear();
      return;
    }

    if (this.steps.length > 1) {
      this.steps.removeAt(index);
    }
  }



  private sanitizeSheetName(raw: string, fallback = 'Lap'): string {
    const invalid = /[\\/?*[\]:]/g;
    let name = (raw ?? '').toString().trim().replace(invalid, ' ').slice(0, 31);
    return name || fallback;
  }
  private ensureUniqueName(used: Set<string>, base: string): string {
    if (!used.has(base)) return base;
    let i = 2;
    while (used.has(`${base} (${i})`)) i++;
    return `${base} (${i})`.slice(0,31);
  }

  exportExcel(): void {
    const wb = new Workbook();
    const used = new Set<string>();

    for (let si = 0; si < this.sheets.length; si++) {
      const sheetGroup = this.sheets.at(si) as FormGroup;
      const stepsFA = sheetGroup.get('steps') as FormArray;
      const rawName = sheetGroup.get('sheetName')?.value ?? '';
      const wsName  = this.ensureUniqueName(used, this.sanitizeSheetName(rawName, `Lap ${si+1}`));
      used.add(wsName);

      const ws = wb.addWorksheet(wsName);
      this.writeSheet(ws, stepsFA);
    }

    wb.xlsx.writeBuffer().then((buffer: ArrayBuffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `${this.testPlanForm.value.title || 'testplan'}.xlsx`);
    });
  }

  private writeSheet(ws: import('exceljs').Worksheet, stepsFA: FormArray): void {
    const alignLeft: Partial<Alignment>   = { vertical: 'top', horizontal: 'left', wrapText: true };
    const alignCenter: Partial<Alignment> = { vertical: 'middle', horizontal: 'center', wrapText: true };
    const headerAlign: Partial<Alignment> = { vertical: 'middle', horizontal: 'left', wrapText: true };
    const thinBorder: Partial<Borders> = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

    ws.columns = [
      { header: '', key: 'A', width: 26 },
      { header: '', key: 'B', width: 28 },
      { header: '', key: 'C', width: 32 },
      { header: '', key: 'D', width: 60 },
      { header: '', key: 'E', width: 30 },
      { header: '', key: 'F', width: 60 },
      { header: '', key: 'G', width: 18 }
    ];

    ws.mergeCells('A1:G1'); ws.getCell('A1').value = 'EESZT - Elektronikus Egészségügyi Szolgáltatás Tér';
    ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true };
    ws.getCell('A1').alignment = alignCenter;

    ws.mergeCells('A2:G2'); ws.getCell('A2').value = 'Tesztelési forgatókönyv';
    ws.getCell('A2').font = { name: 'Calibri', size: 12, bold: true };
    ws.getCell('A2').alignment = alignCenter;

    ws.getRow(1).height = 39; ws.getRow(2).height = 30;
    ws.getRow(4).height = 198.5; ws.getRow(8).height = 30; ws.getRow(12).height = 162.75;

    const metaLabels = ['Honlap:','URL:','Böngésző:','Felhasználó:','Intézmény, szervezeti egység:','Modul:','Shipment:','Redmine:','Rövid leírás:','Előfeltétel:','Futatta:','Készítette:','Dátum','Becsült idő'];
    metaLabels.forEach((label, i) => {
      const r = 3 + i;
      const labelCell = ws.getCell(`A${r}`);
      labelCell.value = label;
      labelCell.font = { name: 'Calibri', size: 10, italic: true };
      labelCell.alignment = alignLeft;
      labelCell.border = thinBorder;

      ws.mergeCells(r, 2, r, 7);
      for (let c = 2; c <= 7; c++) {
        const v = ws.getCell(r, c);
        v.value = '';
        v.font = { name: 'Calibri', size: 10, italic: true };
        v.alignment = alignLeft;
        v.border = thinBorder;
      }
    });

    const headerRowIndex = 18;
    const headers = ['Funkció','Előfeltétel','Feladat/Tesztlépések','Tesztadatok','Elvárt eredmény','✓ / ✘','Megjegyzés'];
    const headerRow = ws.getRow(headerRowIndex);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h; cell.font = { name: 'Calibri', size: 11, bold: true, italic: true };
      cell.alignment = headerAlign; cell.border = thinBorder;
    });
    headerRow.height = 45;

    let rowIdx = headerRowIndex + 1;
    stepsFA.controls.forEach(grp => {
      const v = grp.value || {};
      const row = ws.getRow(rowIdx);

      if (v.isSubtest) {
        ws.mergeCells(rowIdx, 1, rowIdx, 7);
        const c = row.getCell(1);
        const hex = (v.subtestColor || '#555555').toString();
        const bg  = this.hexToARGB(hex);
        const fontColor = this.isDark(hex) ? 'FFFFFFFF' : 'FF000000';
        c.value = String(v.subtestName || '');
        c.font = { name: 'Calibri', size: 11, bold: true, color: { argb: fontColor } };
        c.alignment = alignCenter;
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        for (let col = 1; col <= 7; col++) row.getCell(col).border = thinBorder;
        row.commit(); rowIdx++; return;
      }

      const vals = [
        v.function || '',
        v.precondition || '',
        String(v.stepsText || '').replace(/\r\n/g, '\n'),
        String(v.testData  || '').replace(/\r\n/g, '\n'),
        String(v.expected  || '').replace(/\r\n/g, '\n'),
        '',
        String(v.comment   || '').replace(/\r\n/g, '\n')
      ];

      vals.forEach((val, ci) => {
        const c = row.getCell(ci + 1);
        c.value = val; c.font = { name: 'Calibri', size: 11 };
        c.alignment = alignLeft; c.border = thinBorder;
      });

      row.commit();
      rowIdx++;
    });

    for (let c = 1; c <= 7; c++) {
      ws.getCell(1, c).border = thinBorder;
      ws.getCell(2, c).border = thinBorder;
    }
  }

  private hexToARGB(hex: string): string {
    const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
    if (!m) return 'FF555555';
    return ('FF' + m[1] + m[2] + m[3]).toUpperCase();
  }

  private isDark(hex: string): boolean {
    const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex || '');
    if (!m) return true;
    const r = parseInt(m[1], 16), g = parseInt(m[2], 16), b = parseInt(m[3], 16);
    const yiq = (r*299 + g*587 + b*114) / 1000;
    return yiq < 140;
  }


  async onFileSelected(event: Event): Promise<void> {
    const input = event.target as HTMLInputElement;
    const file  = input.files?.[0];
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      this.ngZone.run(() => {
        this.workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        this.sheetNames = [...this.workbook.SheetNames];
        this.selectedSheetName = this.sheetNames[0] ?? null;
      });

      await Promise.resolve();
      this.importFromSelectedSheet();

    } catch (e) {
      console.error('Excel olvasási hiba:', e);
    } finally {
      input.value = '';
      this.cdr.markForCheck();
    }
  }
  
  onSheetSelected(): void {
    this.importFromSelectedSheet();
  }

  importFromSelectedSheet(): void {
    if (!this.workbook || !this.selectedSheetName) return;
    const sheet = this.workbook.Sheets[this.selectedSheetName];
    if (!sheet) return;

    const aoa: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
    const merges: Array<{ s:{r:number,c:number}, e:{r:number,c:number} }> = (sheet as any)['!merges'] || [];

    const startIdx = Math.max(1, this.startRow) - 1;
    const endIdx   = this.endRow != null && this.endRow >= this.startRow ? this.endRow - 1 : aoa.length - 1;

    const stepsFA = this.activeSheetGroup.get('steps') as FormArray;
    while (stepsFA.length) stepsFA.removeAt(0);

    for (let r = startIdx; r <= endIdx && r < aoa.length; r++) {
      const row = aoa[r] || [];
      if (this.isMergedSubtestRow(r, merges)) { this.addSubtestStep(row); continue; }
      if (this.isNormalStepRow(row)) { this.addNormalStep(row); }
    }
    if (stepsFA.length === 0) stepsFA.push(this.createStep());
  }

  private isMergedSubtestRow(zeroBasedRow: number, merges: Array<{ s: { r: number, c: number }, e: { r: number, c: number } }>): boolean {
    return merges.some(m =>
      m.s.r === zeroBasedRow &&
      m.e.r === zeroBasedRow &&
      m.s.c === 0 &&
      m.e.c === 6
    );
  }

  private isNormalStepRow(row: any[]): boolean {
    const [A,B,C,D,E, ,G] = row;
    return ![A,B,C,D,E,G].every(v => v == null || String(v).trim() === '');
  }

  private addSubtestStep(row: any[]): void {
    const name = row[0] != null ? String(row[0]) : '';
    const fg = this.createStep();
    fg.patchValue({
      isSubtest:   true,
      subtestName: name,
      subtestColor: '#929292',
      function:     '',
      precondition: '',
      stepsText:    '',
      testData:     '',
      expected:     ''
    });
    this.steps.push(fg);
  }

  private addNormalStep(row: any[]): void {
    const [A,B,C,D,E, ,G] = row;
    const fg = this.createStep();
    fg.patchValue({
      isSubtest:   false,
      subtestName: '',
      function:     A != null ? String(A) : '',
      precondition: B != null ? String(B) : '',
      stepsText:    C != null ? String(C).replace(/\r\n/g, '\n') : '',
      testData:     D != null ? String(D).replace(/\r\n/g, '\n') : '',
      expected:     E != null ? String(E).replace(/\r\n/g, '\n') : '',
      comment:      G != null ? String(G).replace(/\r\n/g, '\n') : ''
    });
    this.steps.push(fg);
  }


  clipboard: any[] = [];
  readonly selection = new Set<number>();

  toggleSelect(i: number, checked: boolean): void {
    checked ? this.selection.add(i) : this.selection.delete(i);
  }

  copy(i: number): void {
    const idxs = this.selection.size ? [...this.selection] : [i];
    idxs.sort((a,b)=>a-b);
    this.clipboard = idxs.map(idx => structuredClone(this.steps.at(idx).value));
    this.selection.clear();
  }

  paste(afterIdx: number): void {
    if (!this.clipboard.length) return;

    this.clipboard.forEach((raw, ofs) => {
      const fg = this.createStep();
      fg.patchValue(raw);
      this.steps.insert(afterIdx + 1 + ofs, fg);
    });

    this.clipboard = [];
    this.cdr.detectChanges();
  }

  removeSelected(): void {
    const idxs = [...this.selection].sort((a,b)=>b-a);
    idxs.forEach(i => this.removeStep(i));
    this.selection.clear();
  }

  scrollToStep(idx: number): void {
    const el = document.getElementById(`step-${this.activeSheetIndex}-${idx}`);
    el?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  trackByIndex(_: number, __: unknown): number {
    return _;
  }

}