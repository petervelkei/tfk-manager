export interface TestPlanForm {
  title: string;
  sheets: SheetForm[];
}

export interface SheetForm {
  sheetName: string;
  steps: StepForm[];
}

export interface StepForm {
  isSubtest: boolean;
  subtestName: string;
  subtestColor: string;
  function: string;
  precondition: string;
  stepsText: string;
  testData: string;
  expected: string;
  comment: string;
}