import { Component, EventEmitter, Input, Output } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ReactiveFormsModule, FormGroup } from '@angular/forms';

@Component({
  selector: 'tp-step-card',
  standalone: true,
  imports: [CommonModule, ReactiveFormsModule],
  templateUrl: './step-card.component.html',
  styleUrls: ['./step-card.component.scss']
})
export class StepCardComponent {
  @Input() group!: FormGroup;
  @Input() index!: number;
  @Input() activeSheetIndex!: number;
  @Input() selected = false;
  @Input() clipboardHasContent = false;

  @Output() toggleSelect = new EventEmitter<boolean>();
  @Output() copy = new EventEmitter<void>();
  @Output() paste = new EventEmitter<void>();
  @Output() addAfter = new EventEmitter<void>();
  @Output() remove = new EventEmitter<void>();

  get isSubtest() { return this.group.get('isSubtest')?.value; }
}