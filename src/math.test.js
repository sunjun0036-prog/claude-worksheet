const { add, subtract } = require('./math');

test('add adds two numbers', () => {
  expect(add(2, 3)).toBe(5);
});

test('subtract subtracts two numbers', () => {
  expect(subtract(5, 2)).toBe(3);
});
