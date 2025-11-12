# IT-Support-Ticket-Analysis---Excel
https://github.com/user-attachments/assets/749d4ecd-eef2-4140-aa4e-d50f74f36dd9

This repository contains an end-to-end data analysis project conducted entirely in Microsoft Excel. It analyzes 97,498 IT support tickets over a 5-year period (2016-2020) to evaluate helpdesk performance, measure agent efficiency, and provide actionable recommendations for process improvement.
The final deliverable is a multi-tabbed, dynamic dashboard designed for management to track KPIs and support data-driven decisions.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 📌 Project Overview: 

The primary goal is to analyze IT support operations to identify performance trends, bottlenecks, and areas for improvement.

Key Objectives:

. Evaluate overall helpdesk efficiency by tracking ticket volumes and resolution times.

. Analyze agent performance to identify top performers and pinpoint training needs.

. Uncover trends, seasonal patterns, and workload distribution by category and severity.

. Identify process gaps, such as the critical mismatch between ticket severity and priority.

. Provide data-driven recommendations for automation, staffing, and software upgrades.

. Deliver a dynamic, interactive dashboard for ongoing performance monitoring.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 📂 The Data Analysis is based on two core datasets:

. Tickets Dataset (97,498 records): Contains ticket-level details like Ticket ID, Request Category, Issue Type, Severity, Priority, Resolution Time (Days), Satisfaction Rate, and Date.

. IT Agents Dataset (50 agents): Contains agent-level details like Agent ID, Full Name, Email, and Date of Birth.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 🛠️ Methodology
The project followed a structured data analysis workflow:

# 1. Data Cleaning:

. Corrected over 50,000 inconsistent text entries (e.g., "Mayor" → "Major", "Unassiged" → "Unassigned").
  
.Ensured data integrity by verifying no missing values in critical columns.

# 2. Feature Engineering (Data Transformation):

.Calculated agent Age from Date of Birth using DATE(), TODAY(), and INT() functions.
 
.Extracted email domains using a combination of LEFT, RIGHT, LEN, and FIND functions.

.Created new helper columns for aggregated analysis.
   
# 4. Data Analysis:

.Used Pivot Tables extensively to aggregate data by various dimensions (time, category, agent).

.Merged datasets using VLOOKUP and INDEX/MATCH.

.Performed correlation analysis (CORREL) to investigate the relationship between severity and resolution time.

.Conducted trend analysis to identify daily, monthly, and quarterly patterns.

# 5. Visualization:

.Built a comprehensive, multi-tabbed dashboard.

.Implemented Slicers and Timelines for dynamic, interactive filtering.

.Used various charts (Bar, Line, Pie, Combo) to visualize KPIs.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 📊 Key Insights:
. Process Bottleneck: A critical 82% mismatch rate was found between ticket severity and priority, with 32,000+ 'Normal' tickets incorrectly marked as 'High' priority.

. Category Inefficiency: 'Hardware' (7.63 days) and 'System' (6.62 days) issues are the biggest time consumers, taking 25x longer to resolve than 'Login Access' (0.31 days).

. Stagnant Performance: Despite a 123% growth in ticket volume, the average resolution time remained flat at 4.55 days, indicating a need for process optimization, not just more staff.

. Satisfaction Paradox: Despite stagnant resolution times, the average satisfaction score showed a consistent positive trend, reaching 4.10/5.

. Agent Performance Gaps: Clear variance in agent performance was identified. Top performers include Diana Rojo and Jesus Grajeda, while agents like Lorena and Elena Velez were identified for additional training.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 🚀 Recommendations:

. Upgrade/Reconfigure Software: Implement automation to correctly assign ticket priority based on severity, fixing the 82% mismatch and reducing manual triage.

. Implement Targeted Training: Use performance data to create focused training programs for agents with high resolution times or low satisfaction scores.

. Automate Low-Level Tasks: Automate 'Login Access' (password resets, etc.) to free up agent time for complex issues like 'Hardware' and 'System' tickets.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 🔧 Tools & Techniques

. Primary Tool: Microsoft Excel

. Excel Functions: VLOOKUP, INDEX/MATCH, IF (Array Formulas), DATE, TODAY, INT, LEFT/RIGHT/FIND/LEN, CORREL, AVERAGE.

. Analytical Features: Pivot Tables, Pivot Charts, Slicers, Timelines, Dynamic Dashboards.

. Analytical Concepts: Data Cleaning, Feature Engineering, Trend Analysis, Correlation Analysis, KPI Tracking.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 📎 Repository Contents

. IT_Tickets_Analysis_Dashboard_Excel_File.xlsx: The final interactive Excel dashboard.

. IT_Ticket_Analysis_PPT.pptx: Executive summary presentation of findings and recommendations.

. Task_DOC_File.docx: Detailed Q&A, methodology notes, and calculations.

. Dashboard_Video.mp4: A preview video of the main dashboard.

. Dashboard_Photo: A preview photo of main dashboard.
  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---  ---
# 📬 Author
Sudhanshu Tiwari
