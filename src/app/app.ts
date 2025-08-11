import { Component, signal } from '@angular/core';
import { TestPlanComponent } from './test-plan/test-plan.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [TestPlanComponent],
  templateUrl: './app.html',
  styleUrls: ['./app.scss'] 
})
export class App {
  protected readonly title = signal('tfk-manager');
}