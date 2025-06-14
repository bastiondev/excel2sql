/**
 * Test case for simple query references
 * Tests both direct and iterative query references
 */

export const testName = 'SimpleQuery';

export const queries = {
  data: "SELECT id, name, age from table1;",
  stats: "SELECT count(*) as total, avg(age) as avgAge from table1;"
}

export const queryResults = {
  data: [
    { id: 1, name: 'John', age: 25 },
    { id: 2, name: 'Jane', age: 30 },
    { id: 3, name: 'Bob', age: 35 }
  ],
  stats: [
    { total: 3, avgAge: 30 }
  ]
};
