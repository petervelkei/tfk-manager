import { Component, OnInit, NgZone, ChangeDetectorRef, signal, computed } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormsModule, FormGroup, FormArray, Validators } from '@angular/forms';
import * as XLSX from 'xlsx';

import { SidebarComponent } from './components/sidebar/sidebar.component';
import { SheetTabsComponent } from './components/sheet-tabs/sheet-tabs.component';
import { ImportCardComponent } from './components/import-card/import-card.component';
import { StepCardComponent } from './components/step-card/step-card.component';
import { ExcelService } from './services/excel.service';


@Component({
  selector: 'app-test-plan',
  standalone: true,
  imports: [ CommonModule, ReactiveFormsModule, FormsModule, SidebarComponent, SheetTabsComponent, ImportCardComponent, StepCardComponent ],
  templateUrl: './test-plan.component.html',
  styleUrls: ['./test-plan.component.scss']
})
export class TestPlanComponent implements OnInit {
  testPlanForm!: FormGroup;
  workbook?: XLSX.WorkBook;
  sheetNames: string[] = [];
  selectedSheetName: string | null = null;
  startRow = 19;
  endRow: number | null = null;
  isSidebarOpen = false;
  activeSheetIndex = 0;

  clipboard: any[] = [];
  readonly selection = new Set<number>();

  constructor(
    private readonly fb: FormBuilder,
    private readonly ngZone: NgZone,
    private readonly cdr: ChangeDetectorRef,
    private readonly excel: ExcelService
  ) {}

  ngOnInit(): void {
    this.testPlanForm = this.fb.group({
      title: ['', Validators.required],
      sheets: this.fb.array([ this.createSheet('Tesztelési forgatókönyv') ])
    });
  }

  // --- getters ---
  get sheets(): FormArray { return this.testPlanForm.get('sheets') as FormArray; }
  get activeSheetGroup(): FormGroup { return this.sheets.at(this.activeSheetIndex) as FormGroup; }
  get steps(): FormGroup[] { return (this.activeSheetGroup.get('steps') as FormArray).controls as FormGroup[];}
  get stepsFA(): FormArray { return this.activeSheetGroup.get('steps') as FormArray; }


  // --- sheet ops ---
  addSheet(name?: string): void {
    this.sheets.push(this.createSheet(name ?? `Lap ${this.sheets.length + 1}`));
    this.activeSheetIndex = this.sheets.length - 1;
  }

  removeSheet(idx: number): void {
    if (this.sheets.length <= 1) return;
    this.sheets.removeAt(idx);
    if (this.activeSheetIndex >= this.sheets.length) this.activeSheetIndex = this.sheets.length - 1;
    this.selection.clear();
  }

  switchSheet(idx: number): void {
    this.activeSheetIndex = idx;
    this.selection.clear();
  }

  // --- steps ops ---
  addStep(afterIndex?: number): void {
    const insertAt = ((afterIndex ?? (this.steps.length - 1)) + 1);
    this.stepsFA.insert(insertAt, this.createStep());
  }

  removeStep(index: number): void {
    if (this.selection.size) {
      const stepsFA = this.activeSheetGroup.get('steps') as FormArray;
      [...this.selection].sort((a, b) => b - a).forEach(i => { if (stepsFA.length > 1) stepsFA.removeAt(i); });
      this.selection.clear();
      return;
    }
    if (this.steps.length > 1) {
      const stepsFA = this.activeSheetGroup.get('steps') as FormArray;
      stepsFA.removeAt(index);
    }
  }

  toggleSelect(i: number, checked: boolean): void { checked ? this.selection.add(i) : this.selection.delete(i); }
  
  copy(i: number): void {
    const idxs = this.selection.size ? [...this.selection] : [i];
    idxs.sort((a,b)=>a-b);
    this.clipboard = idxs
      .map(idx => {
        const step = this.stepsFA.at(idx) as FormGroup | null;
        return step ? structuredClone(step.value) : null;
      })
      .filter((v): v is object => v !== null);
    this.selection.clear();
  }

  paste(afterIdx: number): void {
    if (!this.clipboard.length) return;
    const stepsFA = this.activeSheetGroup.get('steps') as FormArray;
    this.clipboard.forEach((raw, ofs) => { const fg = this.createStep(); fg.patchValue(raw); stepsFA.insert(afterIdx + 1 + ofs, fg); });
    this.clipboard = []; this.cdr.detectChanges();
  }

  scrollToStep(idx: number): void { document.getElementById(`step-${this.activeSheetIndex}-${idx}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }

  // --- import ---
  async onFileSelected(file: File): Promise<void> {
    try {
      const buffer = await file.arrayBuffer();
      this.ngZone.run(() => {
        this.workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        this.sheetNames = [...this.workbook!.SheetNames];
        this.selectedSheetName = this.sheetNames[0] ?? null;
      });
      await Promise.resolve();
      this.importFromSelectedSheet();
    } catch (e) {
      console.error('Excel olvasási hiba:', e);
    } finally {
      this.cdr.markForCheck();
    }
  }

  onSheetSelected(name: string) { this.selectedSheetName = name; this.importFromSelectedSheet(); }
  
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
    return merges.some(m => m.s.r === zeroBasedRow && m.e.r === zeroBasedRow && m.s.c === 0 && m.e.c === 6);
  }

  private isNormalStepRow(row: any[]): boolean {
    const [A,B,C,D,E, ,G] = row; return ![A,B,C,D,E,G].every(v => v == null || String(v).trim() === '');
  }

  private addSubtestStep(row: any[]): void {
    const name = row[0] != null ? String(row[0]) : '';
    const fg = this.createStep();
    fg.patchValue({
      isSubtest: true,
      subtestName: name,
      subtestColor: '#929292',
      function: '', precondition: '', stepsText: '', testData: '', expected: ''
    });
    this.stepsFA.push(fg);
  }

  private addNormalStep(row: any[]): void {
    const [A,B,C,D,E, ,G] = row;
    const fg = this.createStep();
    fg.patchValue({
      isSubtest: false,
      subtestName: '',
      function:     A ? String(A) : '',
      precondition: B ? String(B) : '',
      stepsText:    C ? String(C).replace(/\r\n/g, '\n') : '',
      testData:     D ? String(D).replace(/\r\n/g, '\n') : '',
      expected:     E ? String(E).replace(/\r\n/g, '\n') : '',
      comment:      G ? String(G).replace(/\r\n/g, '\n') : ''
    });
    this.stepsFA.push(fg);
  }

  // --- export ---
  exportExcel(): void { this.excel.export(this.testPlanForm.value); }

  // --- factories ---
  private createSheet(name = 'Lap 1'): FormGroup {
    return this.fb.group({ sheetName: [name, Validators.required], steps: this.fb.array([ this.createStep() ]) });
  }

  private createStep(): FormGroup {
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
}