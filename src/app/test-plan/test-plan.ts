import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule, FormBuilder, FormGroup, FormArray, Validators } from '@angular/forms';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';

@Component({
  standalone: true,
  selector: 'app-test-plan',
  imports: [ CommonModule, ReactiveFormsModule ],
  templateUrl: './test-plan.html',
  styleUrls: ['./test-plan.scss']
})
export class TestPlan implements OnInit {
  testPlanForm!: FormGroup;

  constructor(private readonly fb: FormBuilder) {}

  ngOnInit(): void {
    this.testPlanForm = this.fb.group({
      title: ['', Validators.required],
      steps: this.fb.array([ this.createStep() ])
    });
  }

  createStep(): FormGroup {
    return this.fb.group({
      function:     ['', Validators.required],
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
    const controls = this.steps.controls;
    const aoa: any[][] = [];
    aoa.push(['EESZT - Elektronikus Egészségügyi Szolgáltatás Tér']);
    aoa.push(['Tesztelési forgatókönyv']);

    const metaLabels = [
      'Honlap:', 'URL:', 'Böngésző:', 'Felhasználó:', 'Intézmény, szervezeti egység:',
      'Modul:', 'Shipment:', 'Redmine:', 'Rövid leírás:', 'Előfeltétel:',
      'Futatta:', 'Készítette:', 'Dátum', 'Becsült idő'
    ];

    metaLabels.forEach(label => aoa.push([label, '']));

    aoa.push([]);

    const tableHeaders = [
      'Funkció',
      'Előfeltétel',
      'Feladat/Tesztlépések',
      'Tesztadatok',
      'Elvárt eredmény',
      '✔︎ / ✘',
      'Megjegyzés'
    ];
    aoa.push(tableHeaders);

    controls.forEach(grp => {
      const v = grp.value || {};
      aoa.push([
        v.function || '',
        v.precondition || '',
        (v.stepsText || '').replace(/\r\n/g, '\n'),
        (v.testData || '').replace(/\r\n/g, '\n'),
        (v.expected || '').replace(/\r\n/g, '\n'),
        '',
        ''
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const merges: XLSX.Range[] = [];

    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } });
    merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 6 } });

    const metaStart = 2;
    for (let i = 0; i < metaLabels.length; i++) {
      merges.push({ s: { r: metaStart + i, c: 1 }, e: { r: metaStart + i, c: 6 } });
    }

    (ws as any)['!merges'] = merges;

    (ws as any)['!cols'] = [
      { wch: 26 },
      { wch: 28 },
      { wch: 32 },
      { wch: 60 },
      { wch: 30 },
      { wch: 60 },
      { wch: 18 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Tesztforgatókönyv');

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    saveAs(blob, `${this.testPlanForm.value.title || 'testplan'}.xlsx`);
  }
}