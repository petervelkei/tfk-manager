import { Component, Input, Output, EventEmitter } from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'tp-sidebar',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './sidebar.component.html',
  styleUrls: ['./sidebar.component.scss']
})
export class SidebarComponent {
  @Input() stepsCount = 0;
  @Input() activeSheetIndex = 0;
  @Input() selectedIndices = new Set<number>();
  @Input() open = false;

  @Output() toggleOpen = new EventEmitter<void>();
  @Output() scrollTo = new EventEmitter<number>();

  trackByIndex(i: number) { return i; }
}