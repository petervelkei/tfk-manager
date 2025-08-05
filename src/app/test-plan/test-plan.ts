import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import {
  ReactiveFormsModule, FormBuilder, FormsModule,
  FormGroup, FormArray, Validators
} from '@angular/forms';
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

  constructor(private readonly fb: FormBuilder) {}

  ngOnInit(): void {
    this.testPlanForm = this.fb.group({
      title: ['', Validators.required],
      steps: this.fb.array([ this.createStep() ])
    });
  }

  createStep(): FormGroup {
    return this.fb.group({
      isSubtest:    [false],
      subtestName:  [''],
      function:     [''],
      precondition: [''],
      testData:     [''],
      stepsText:    [''],
      expected:     ['']
    });
  }

  get steps(): FormArray {
    return this.testPlanForm.get('steps') as FormArray;
  }

  addStep(): void {
    this.steps.push(this.createStep());
  }

  removeStep(index: number): void {
    if (this.steps.length > 1) {
      this.steps.removeAt(index);
    }
  }


  exportExcel(): void {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Tesztforgatókönyv');

    const alignLeft: Partial<Alignment>   = { vertical: 'top',    horizontal: 'left',  wrapText: true };
    const alignCenter: Partial<Alignment> = { vertical: 'middle', horizontal: 'center', wrapText: true };
    const headerAlign: Partial<Alignment> = { vertical: 'middle', horizontal: 'left',   wrapText: true };

    const thinBorder: Partial<Borders> = {
      top:    { style: 'thin' },
      left:   { style: 'thin' },
      bottom: { style: 'thin' },
      right:  { style: 'thin' }
    };

    ws.columns = [
      { header: '', key: 'A', width: 26 },
      { header: '', key: 'B', width: 28 },
      { header: '', key: 'C', width: 32 },
      { header: '', key: 'D', width: 60 },
      { header: '', key: 'E', width: 30 },
      { header: '', key: 'F', width: 60 },
      { header: '', key: 'G', width: 18 }
    ];

    ws.mergeCells('A1:G1');
    ws.getCell('A1').value = 'EESZT - Elektronikus Egészségügyi Szolgáltatás Tér';
    ws.getCell('A1').font = { name: 'Calibri', size: 14, bold: true };
    ws.getCell('A1').alignment = alignCenter;

    ws.mergeCells('A2:G2');
    ws.getCell('A2').value = 'Tesztelési forgatókönyv';
    ws.getCell('A2').font = { name: 'Calibri', size: 12, bold: true };
    ws.getCell('A2').alignment = alignCenter;

    ws.getRow(1).height  = 39;
    ws.getRow(2).height  = 30;
    ws.getRow(4).height  = 198.5;
    ws.getRow(8).height  = 30;
    ws.getRow(12).height = 162.75;

    const metaLabels = [
      'Honlap:', 'URL:', 'Böngésző:', 'Felhasználó:', 'Intézmény, szervezeti egység:',
      'Modul:', 'Shipment:', 'Redmine:', 'Rövid leírás:', 'Előfeltétel:',
      'Futatta:', 'Készítette:', 'Dátum', 'Becsült idő'
    ];
    metaLabels.forEach((label, i) => {
      const r = 3 + i;
      const labelCell = ws.getCell(`A${r}`);
      labelCell.value = label;
      labelCell.font = { name: 'Calibri', size: 10, italic: true };
      labelCell.alignment = alignLeft;
      labelCell.border = thinBorder;

      ws.mergeCells(r, 2, r, 7);
      const valueCell = ws.getCell(r, 2);
      valueCell.value = '';
      valueCell.font = { name: 'Calibri', size: 10, italic: true };
      valueCell.alignment = alignLeft;
      for (let c = 2; c <= 7; c++) {
        ws.getCell(r, c).border = thinBorder;
      }
    });

    const headerRowIndex = 18;
    const headers = [
      'Funkció',
      'Előfeltétel',
      'Feladat/Tesztlépések',
      'Tesztadatok',
      'Elvárt eredmény',
      '✓ / ✘',
      'Megjegyzés'
    ];
    const headerRow = ws.getRow(headerRowIndex);
    headers.forEach((h, i) => {
      const cell = headerRow.getCell(i + 1);
      cell.value = h;
      cell.font = { name: 'Calibri', size: 11, bold: true, italic: true };
      cell.alignment = headerAlign;
      cell.border = thinBorder;
    });
    headerRow.height = 45;

    let rowIdx = headerRowIndex + 1;
    this.steps.controls.forEach(grp => {
      const v = grp.value || {};
      const row = ws.getRow(rowIdx);

      if (v.isSubtest) {
        ws.mergeCells(rowIdx, 1, rowIdx, 7);
        const c = row.getCell(1);
        c.value = String(v.subtestName || '');
        c.font = { name: 'Calibri', size: 11, bold: true };
        c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };

        for (let col = 1; col <= 7; col++) {
          row.getCell(col).border = {
            top:    { style: 'thin' },
            left:   { style: 'thin' },
            bottom: { style: 'thin' },
            right:  { style: 'thin' }
          };
        }

        row.height = 28;
        row.commit();
        rowIdx++;
        return;
      }

      const vals = [
        v.function || '',
        v.precondition || '',
        String(v.stepsText || '').replace(/\r\n/g, '\n'),
        String(v.testData  || '').replace(/\r\n/g, '\n'),
        String(v.expected  || '').replace(/\r\n/g, '\n'),
        '',
        ''
      ];

      vals.forEach((val, ci) => {
        const c = row.getCell(ci + 1);
        c.value = val;
        c.font = { name: 'Calibri', size: 11 };
        c.alignment = alignLeft;
        c.border = thinBorder;
      });

      row.height = 25;
      row.commit();
      rowIdx++;
    });

    for (let c = 1; c <= 7; c++) {
      ws.getCell(1, c).border = thinBorder;
      ws.getCell(2, c).border = thinBorder;
    }

    wb.xlsx.writeBuffer().then((buffer: ArrayBuffer) => {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      saveAs(blob, `${this.testPlanForm.value.title || 'testplan'}.xlsx`);
    });
  }


  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = new Uint8Array(reader.result as ArrayBuffer);
        this.workbook = XLSX.read(data, { type: 'array' });

        this.sheetNames = this.workbook.SheetNames.slice();
        this.selectedSheetName = this.sheetNames[0] || null;

        if (this.selectedSheetName) {
          this.importFromSelectedSheet();
        }
      } catch (e) {
        console.error('Excel olvasási hiba:', e);
      } finally {
        input.value = '';
      }
    };
    reader.onerror = (err) => {
      console.error('FileReader hiba:', err);
    };
    reader.readAsArrayBuffer(file);
  }

  onSheetSelected(): void {
    this.importFromSelectedSheet();
  }

  importFromSelectedSheet(): void {
    if (!this.workbook || !this.selectedSheetName) return;

    const sheet = this.workbook.Sheets[this.selectedSheetName];
    if (!sheet) return;

    const aoa: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
    const merges: Array<{ s: { r: number, c: number }, e: { r: number, c: number } }> =
      (sheet as any)['!merges'] || [];

    const startIdx = Math.max(1, this.startRow) - 1;

    while (this.steps.length) this.steps.removeAt(0);

    for (let r = startIdx; r < aoa.length; r++) {
      const row = aoa[r] || [];
      if (this.isMergedSubtestRow(r, merges)) {
        this.addSubtestStep(row);
        continue;
      }
      if (this.isNormalStepRow(row)) {
        this.addNormalStep(row);
      }
    }

    if (this.steps.length === 0) {
      this.steps.push(this.createStep());
    }
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
    const [A, B, C, D, E] = row;
    return ![A, B, C, D, E].every(v => v == null || String(v).trim() === '');
  }

  private addSubtestStep(row: any[]): void {
    const name = row[0] != null ? String(row[0]) : '';
    const fg = this.createStep();
    fg.patchValue({
      isSubtest:   true,
      subtestName: name,
      function:     '',
      precondition: '',
      stepsText:    '',
      testData:     '',
      expected:     ''
    });
    this.steps.push(fg);
  }

  private addNormalStep(row: any[]): void {
    const [A, B, C, D, E] = row;
    const fg = this.createStep();
    fg.patchValue({
      isSubtest:   false,
      subtestName: '',
      function:     A != null ? String(A) : '',
      precondition: B != null ? String(B) : '',
      stepsText:    C != null ? String(C).replace(/\r\n/g, '\n') : '',
      testData:     D != null ? String(D).replace(/\r\n/g, '\n') : '',
      expected:     E != null ? String(E).replace(/\r\n/g, '\n') : ''
    });
    this.steps.push(fg);
  }
}