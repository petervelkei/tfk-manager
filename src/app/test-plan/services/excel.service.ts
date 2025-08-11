import { Injectable } from '@angular/core';
import type { TestPlanForm, SheetForm, StepForm } from '../models/test-plan.models';
import { Workbook, Alignment, Borders } from 'exceljs';
import { saveAs } from 'file-saver';

@Injectable({ providedIn: 'root' })
export class ExcelService {
	
  export(form: TestPlanForm): void {
    const wb = new Workbook();
    const used = new Set<string>();

    form.sheets.forEach((sheet, si) => {
      const wsName = this.ensureUniqueName(used, this.sanitizeSheetName(sheet.sheetName || `Lap ${si + 1}`));
      used.add(wsName);
      const ws = wb.addWorksheet(wsName);
      this.writeSheet(ws, sheet.steps);
    });

    wb.xlsx.writeBuffer().then((buffer: ArrayBuffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `${form.title || 'testplan'}.xlsx`);
    });
  }

  private writeSheet(ws: import('exceljs').Worksheet, steps: StepForm[]): void {
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
    steps.forEach(v => {
      const row = ws.getRow(rowIdx);

      if (v.isSubtest) {
        ws.mergeCells(rowIdx, 1, rowIdx, 7);
        const c = row.getCell(1);
        const hex = v.subtestColor || '#555555';
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

  private sanitizeSheetName(raw: string, fallback = 'Lap'): string {
    const invalid = /[\\\/?*\[\]:]/g;
    const name = (raw ?? '').toString().trim().replace(invalid, ' ').slice(0, 31);
    return name || fallback;
  }

  private ensureUniqueName(used: Set<string>, base: string): string {
    if (!used.has(base)) return base;
    let i = 2; while (used.has(`${base} (${i})`)) i++;
    return `${base} (${i})`.slice(0,31);
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
}