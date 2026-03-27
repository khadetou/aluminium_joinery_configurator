from __future__ import annotations

import ast


class FormulaError(ValueError):
    """Raised when a configured formula cannot be evaluated safely."""


ALLOWED_BINOPS = {
    ast.Add: lambda a, b: a + b,
    ast.Sub: lambda a, b: a - b,
    ast.Mult: lambda a, b: a * b,
    ast.Div: lambda a, b: a / b,
    ast.Pow: lambda a, b: a ** b,
    ast.Mod: lambda a, b: a % b,
}

ALLOWED_UNARYOPS = {
    ast.UAdd: lambda a: +a,
    ast.USub: lambda a: -a,
}


def safe_eval_formula(expression: str, variables: dict[str, float]) -> float:
    """Evaluate a limited arithmetic expression for configured rules."""
    if not expression:
        raise FormulaError("Missing expression.")
    try:
        node = ast.parse(expression, mode="eval")
    except SyntaxError as exc:
        raise FormulaError(str(exc)) from exc
    return float(_eval_node(node.body, variables))


def _eval_node(node, variables: dict[str, float]) -> float:
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return float(node.value)
    if isinstance(node, ast.Name):
        if node.id not in variables:
            raise FormulaError(f"Unknown variable '{node.id}'.")
        return float(variables[node.id])
    if isinstance(node, ast.BinOp) and type(node.op) in ALLOWED_BINOPS:
        return ALLOWED_BINOPS[type(node.op)](
            _eval_node(node.left, variables),
            _eval_node(node.right, variables),
        )
    if isinstance(node, ast.UnaryOp) and type(node.op) in ALLOWED_UNARYOPS:
        return ALLOWED_UNARYOPS[type(node.op)](_eval_node(node.operand, variables))
    raise FormulaError(f"Unsupported expression element: {type(node).__name__}")

