# Project: Yield Curve Modeling in Excel

In this project, I developed a comprehensive Yield Curve Modeling tool in Microsoft Excel **(using Daily Treasury Par Yield Curve Rates)** to estimate and visualize interest rate term structures. The objective was to fit smooth yield curves to observed market data for better pricing, risk management, and financial analysis.

# Official Source Link:
https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202505

#  Key Objectives:
- To model and fit the term structure of interest rates with high precision
- To compare different methods of curve fitting and evaluate their effectiveness
- To implement each model in Excel with user-friendly features for customization and interpretation

# Modeling Techniques Used:

**1.  Nelson-Siegel Model:**
- Parametric model with 4 parameters: β1, β2, β3, and τ
- Captures level, slope, and curvature of the yield curve
- Used Solver in Excel to calibrate parameters by minimizing the error between observed and fitted yields

**2. Nelson-Siegel-Svensson Model:**
- Extension of Nelson-Siegel with 6 parameters: β1, β2, β3, β4, τ₁, τ₂
- Provides additional flexibility to model the hump or dip in long-term yields
- Solver-based optimization to calibrate the full model and compare its performance to the standard Nelson-Siegel

**3. Linear Interpolation:**
- Constructed the yield curve by connecting known yields at specific maturities with straight lines
- Simple, fast, and useful for short-term estimations
- Implemented using Excel’s built-in functions with dynamic range inputs

# Visualization:
- Created dynamic charts to compare fitted curves across the three models
- Included forward rate curves, spot rate curves, and error plots
- Enabled interactive model selection and parameter sensitivity analysis

# Outcomes & Insights:
- Nelson-Siegel-Svensson consistently provided a better fit for complex yield structures
- Linear interpolation was most effective for short-term, simple applications
- Gained strong hands-on experience in financial modeling, curve fitting, and Excel-based optimization techniques

# Tools & Skills Used:
- Microsoft Excel (Advanced Formulas, Solver, Dynamic Charts)
- Yield Curve Theory & Parametric Modeling
- Optimization and Least-Squares Fitting Techniques

# Takeaway:
This project enhanced my quantitative finance and Excel modeling skills, particularly in the domain of interest rate modeling. It also reinforced the importance of model selection in financial decision-making, portfolio valuation, and risk management.
