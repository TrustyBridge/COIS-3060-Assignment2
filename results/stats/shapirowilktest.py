from scipy import stats

aop_complexity = [2.8, 3.4, 2.9, 2.4, 3.1]
oop_complexity = [2.2, 2.6, 2.1, 3.2, 2.0]

aop_churn = [261, 218, 308, 388, 119]
oop_churn = [142, 318, 96, 228, 114]

stat, p = stats.shapiro(aop_complexity)
print(f"AOP Complexity: W={stat:.4f}, p={p:.4f}")

stat, p = stats.shapiro(oop_complexity)
print(f"OOP Complexity: W={stat:.4f}, p={p:.4f}")

stat, p = stats.shapiro(aop_churn)
print(f"AOP Churn: W={stat:.4f}, p={p:.4f}")

stat, p = stats.shapiro(oop_churn)
print(f"OOP Churn: W={stat:.4f}, p={p:.4f}")