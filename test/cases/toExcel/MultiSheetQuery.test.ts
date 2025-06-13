export const testName = 'MultiSheetQuery';

export const queryResults = {
  employees: [
    { id: 1, name: 'John', department: 'Sales', salary: 50000 },
    { id: 2, name: 'Jane', department: 'Engineering', salary: 80000 },
    { id: 3, name: 'Bob', department: 'Sales', salary: 45000 }
  ],
  departments: [
    { name: 'Sales', headcount: 2, budget: 100000 },
    { name: 'Engineering', headcount: 1, budget: 150000 }
  ],
  summary: [
    { total_employees: 3, total_budget: 250000 }
  ]
};
