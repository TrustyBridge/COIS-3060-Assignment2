
from cliffs_delta import cliffs_delta

aop_complexity = [2.8, 3.4, 2.9, 2.4, 3.1]
oop_complexity = [2.2, 2.6, 2.1, 3.2, 2.0]

aop_churn = [261, 218, 308, 388, 119]
oop_churn = [142, 318, 96, 228, 114]

d, res = cliffs_delta(aop_complexity, oop_complexity)
print(f"Complexity — Cliff's Delta: {d:.4f} ({res})")

d, res = cliffs_delta(aop_churn, oop_churn)
print(f"Churn — Cliff's Delta: {d:.4f} ({res})")