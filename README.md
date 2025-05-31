# Project: Yield Curve Modeling in Excel
In this project, I developed a comprehensive Yield Curve Modeling tool in Microsoft Excel **(using Daily Treasury Par Yield Curve Rates)** to estimate and visualize interest rate term structures. The objective was to fit smooth yield curves to observed market data for better pricing, risk management, and financial analysis.

# Official Source Link:
https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202505

# Tools Used: Microsoft Excel
# Calibration Method: Ordinary Least Squares (OLS)

#  Key Objectives:
- To model and fit the term structure of interest rates with high precision
- To compare different methods of curve fitting and evaluate their effectiveness
- To implement each model in Excel with user-friendly features for customization and interpretation

# Modeling Techniques Used:

**Nelson-Siegel Model (NS):**
A widely used parametric model to fit the yield curve using four parameters (β1, β2, β3, τ).
- Captures level, slope, and curvature components
- Ideal for smooth yield curves
- Implemented using non-linear curve fitting via OLS to minimize the sum of squared errors between model and market yields

**2. Nelson-Siegel-Svensson Model (NSS):**
An extension of the NS model that adds a second curvature term for better flexibility.
- Utilizes six parameters for enhanced fit (β1, β2, β3, β4, τ₁, τ₂)
- Especially useful for fitting more complex yield curve shapes
- Fitted using OLS optimization through Excel Solver

**3. Linear Interpolation:**
A non-parametric method that connects observed spot or zero-coupon rates linearly between known maturities.
- Constructed the yield curve by connecting known yields at specific maturities with straight lines
- Simple, fast, and useful for short-term estimations
- Implemented using Excel’s built-in functions with dynamic range inputs

# Calibration & Implementation:
- **Data Input**: A range of maturities and corresponding observed yields.
- **OLS Calibration**: All model parameters were estimated using Ordinary Least Squares (OLS) by minimizing the error between model-implied and actual yields.
- **Excel Solver**: Used to optimize model parameters with constraints to ensure realistic yield curve behavior.

# Visualization:
- Created dynamic charts to compare fitted curves across the three models
- Included forward rate curves, spot rate curves, and error plots
- Enabled interactive model selection and parameter sensitivity analysis

# Outcomes & Insights:
- Observed that **NSS provided the best fit** for complex yield curve shapes, followed by **NS** and then **Linear Interpolation**.
- Gained deep insight into fixed-income analytics, interest rate modeling, and optimization in Excel without relying on specialized financial software.
- Demonstrated the power of Excel for financial modeling, especially for academic, research, or preliminary prototyping purposes.

# Tools & Skills Used:
- Microsoft Excel (Advanced Formulas, Solver, Dynamic Charts)
- Yield Curve Theory & Parametric Modeling
- Optimization and Least-Squares Fitting Techniques

# Takeaway:
This project enhanced my quantitative finance and Excel modeling skills, particularly in the domain of interest rate modeling. It also reinforced the importance of model selection in financial decision-making, portfolio valuation, and risk management.
