import { Component, signal } from '@angular/core';
import { TestPlan } from './test-plan/test-plan';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [TestPlan],
  templateUrl: './app.html',
  styleUrls: ['./app.scss']
})
export class App {
  protected readonly title = signal('tfk-manager');
}