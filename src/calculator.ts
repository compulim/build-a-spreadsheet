/**
 * A very simple calculator supporting `+`, `-`, `*`, `/`, unary signs and nested parentheses.
 *
 * Returns the computed number, or `undefined` if the expression is invalid or the result is not finite.
 */
export function calculate(expression: string): number | undefined {
  const tokens = expression.match(/\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+(?:[eE][-+]?\d+)?|[-+*/()]|\S/gu) || [];
  let index = 0;

  const peek = () => tokens[index];

  const parseExpression = (): number => {
    let value = parseTerm();

    while (peek() === '+' || peek() === '-') {
      const operator = tokens[index++];
      const right = parseTerm();

      value = operator === '+' ? value + right : value - right;
    }

    return value;
  };

  const parseTerm = (): number => {
    let value = parseFactor();

    while (peek() === '*' || peek() === '/') {
      const operator = tokens[index++];
      const right = parseFactor();

      value = operator === '*' ? value * right : value / right;
    }

    return value;
  };

  const parseFactor = (): number => {
    const token = tokens[index++];

    if (token === '+' || token === '-') {
      const value = parseFactor();

      return token === '-' ? -value : value;
    } else if (token === '(') {
      const value = parseExpression();

      if (tokens[index++] !== ')') {
        throw new Error('Missing closing parenthesis.');
      }

      return value;
    } else if (token !== undefined && /^(?:\d|\.\d)/u.test(token)) {
      return Number(token);
    }

    throw new Error('Unexpected token.');
  };

  try {
    const value = parseExpression();

    return index === tokens.length && Number.isFinite(value) ? value : undefined;
  } catch {
    return undefined;
  }
}
