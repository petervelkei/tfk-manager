import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormArray } from '@angular/forms';

@Component({
  selector: 'tp-sheet-tabs',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './sheet-tabs.component.html',
  styleUrls: ['./sheet-tabs.component.scss']
})
export class SheetTabsComponent {
  @Input() sheets!: FormArray;
  @Input() activeIndex = 0;
  @Output() switch = new EventEmitter<number>();
  @Output() add = new EventEmitter<void>();
  @Output() remove = new EventEmitter<number>();
}