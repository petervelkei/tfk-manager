import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'tp-import-card',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './import-card.component.html',
  styleUrls: ['./import-card.component.scss']
})
export class ImportCardComponent {
  @Input() sheetNames: string[] = [];
  @Input() selectedSheetName: string | null = null;
  @Input() startRow = 19;
  @Input() endRow: number | null = null;

  @Output() fileSelected = new EventEmitter<File>();
  @Output() sheetChange = new EventEmitter<string>();
  @Output() startRowChange = new EventEmitter<number>();
  @Output() endRowChange = new EventEmitter<number | null>();
  @Output() importClick = new EventEmitter<void>();

  onFileInput(e: Event) {
    const f = (e.target as HTMLInputElement).files?.[0];
    if (f) this.fileSelected.emit(f);
    (e.target as HTMLInputElement).value = '';
  }
}