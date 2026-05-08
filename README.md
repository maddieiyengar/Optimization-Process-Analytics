# Optimization & Process Analytics (BIA 650)

---

## Repository Overview
This repository features a collection of financial optimization models designed for asset management and corporate resource allocation. By leveraging **Microsoft Excel Solver** and **VBA**, these projects apply linear, integer, and quadratic programming to solve high-level business constraints and drive operational efficiency.

---

## Project Descriptions

### 1. Hedge Fund Optimization
* **Goal**: Maximize risk-adjusted returns through **Mean-Variance Optimization (MVO)**.
* **Description**: This model identifies the **Efficient Frontier** for a diverse asset class mix. It optimizes portfolio weights to achieve the highest possible return for a specific risk tolerance, incorporating constraints such as minimum return thresholds and volatility caps.
* **Key Techniques**: Quadratic Programming, Sensitivity Analysis, and Portfolio Variance Minimization.

---

### 2. Strategic Capital Budgeting
* **Goal**: Optimize project selection to maximize total **Net Present Value (NPV)**.
* **Description**: A strategic framework designed to select the most profitable combination of investment projects under strict financial limitations. It successfully manages multi-period budget caps, project interdependencies, and liquidity requirements.
* **Key Techniques**: **Binary Integer Programming**, Resource Allocation, and Constraint Management.

---

### 3. Index Replication Model
* **Goal**: Minimize tracking error against a target benchmark index.
* **Description**: This model replicates the performance of a market index using a curated subset of assets, significantly reducing transaction costs while maintaining market exposure. It includes a deep dive into how market fluctuations impact portfolio stability.
* **Key Techniques**: **Tracking Error Minimization**, Shadow Pricing, and Sensitivity Reporting.

---

## Tools & Technologies

| Category | Details |
| :--- | :--- |
| **Software** | Microsoft Excel (Solver Add-in) |
| **Analysis** | Sensitivity Reporting, Heuristic Search, and Out-of-Sample Validation |
| **Automation** | **VBA-enhanced logic** for iterative model processing |

---

## Instructions for Use
1. **Enable Solver**: Ensure the **Excel Solver Add-in** is enabled via **File > Options > Add-ins > Excel Add-ins > Go**.
2. **Enable Macros**: Ensure **Macros** are enabled for workbooks requiring VBA-driven iterative processing.
3. **Run Models**: Access individual `.xlsm` files to review specific constraint configurations, objective functions, and sensitivity reports.

