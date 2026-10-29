from pathlib import Path
from pptx import Presentation
import pandas as pd
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

BASE_DIR= Path(__file__).resolve().parent

DATA_FILE=BASE_DIR/"novaretail_dataset.xlsx"
TEMPLATE_FILE=BASE_DIR/"novaretail_template.pptx"
OUTPUT_FILE=BASE_DIR/"novaretail_quarterly_review.pptx"

SHEETS =["Regional_Performance","Customer_Experience","Operational_Risk"]
# -----------------------------------------------------
# Common styling for Slides 2, 3 and 4
# -----------------------------------------------------

CONTENT_TITLE_SIZE = 24
CONTENT_INSIGHT_SIZE = 13
CONTENT_ICON_SIZE = 0.90

CONTENT_TITLE_COLOR = RGBColor(31, 45, 91)
CONTENT_TEXT_COLOR = RGBColor(0, 0, 0)

def load_inputs():
    for path in (DATA_FILE,TEMPLATE_FILE):
        if not path.is_file():
            raise FileNotFoundError(f"Required file not found:{path}")

    workbook = pd.ExcelFile(DATA_FILE)
    missing_sheets = set(SHEETS)-set(workbook.sheet_names)

    if missing_sheets:
        raise ValueError(f"Workbook is missing sheets:{sorted(missing_sheets)}")

    data={sheet:pd.read_excel(workbook,sheet_name=sheet)
          for sheet in SHEETS}

    presentation = Presentation(TEMPLATE_FILE)

    if len(presentation.slides)!=5:
        raise ValueError(f"Expected 5 template slides,found{len(presentation.slides)}")

    for sheet_name, df in data.items():
        if df.empty:
            raise ValueError(f"Sheet'{sheet_name}' contains no data.")
        df.columns=df.columns.astype(str).str.strip()

    return data,presentation

def set_text(shape,text):
    if not shape.has_text_frame:
        return
    shape.text_frame.clear()
    paragraph = shape.text_frame.paragraphs[0]
    paragraph.text=str(text)
    paragraph.font.size=Pt(12)

def set_slide1_text_color(slide,shape_name,color = RGBColor(255,255,255)):
    shape = find_shape(slide,shape_name)
    if shape is None or not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb=color
            run.font.bold = True

def find_shape(slide,name):
    for shape in slide.shapes:
        if shape.name== name: 
            return shape
    return None

def fill_shape(slide,name,text):
    shape= find_shape(slide,name)
    if shape is not None:
        set_text(shape,text)

def quarter_order(value):
    import re

    text = str(value).strip().upper()

    # Q2-2026 / Q2 2026
    match = re.search(
        r"Q([1-4])\s*[- ]?\s*(20\d{2})",
        text
    )

    if match:
        quarter = int(match.group(1))
        year = int(match.group(2))

        return year, quarter

    # 2026-Q2 / 2026 Q2
    match = re.search(
        r"(20\d{2})\s*[- ]?\s*Q([1-4])",
        text
    )

    if match:
        year = int(match.group(1))
        quarter = int(match.group(2))

        return year, quarter

    return 0, 0

def ordered_quarters(df):
    quarters = df["Quarter"].dropna().unique().tolist()
    return sorted(quarters,key=quarter_order)

def aggregate(df,quarter):
    return df[df["Quarter"]== quarter].copy()

def change_pct(current,previous):
    if pd.isna(previous) or previous ==0:
        return None
    return (current-previous)/abs(previous)*100
def fmt(value,suffix=""):
    if pd.isna(value):
        return "N/A"
    return f"{value:,.2f}{suffix}"

def set_chart(slide,shape_name,df,metric,title):
    shape=find_shape(slide,shape_name)
    if shape is None:
        return
        
    if isinstance(metric,list):
        metric = metric[0]

    if metric not in df.columns:
        print(f"Missing chart metric: {metric}")
        return
    regional = df.groupby("Region",as_index=False)[metric].mean()

    if regional.empty:
        return
    chart_data=CategoryChartData()
    chart_data.categories= regional["Region"].astype(str).tolist()
    chart_data.add_series(
    str(metric),
    regional[metric]
    .fillna(0)
    .astype(float)
    .tolist()
    )

    left,top= shape.left,shape.top
    width,height = shape.width,shape.height
    shape._element.getparent().remove(shape._element)

    chart= slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left,top,width,height,chart_data).chart

    chart.has_title= True
    chart.chart_title.text_frame.text = title
    chart.has_legend = False
    #chart.value_axis.has_major_gridlines = True

def remove_shape(slide, shape_name):
    """Remove a named shape from a slide."""

    shape = find_shape(slide, shape_name)

    if shape is None:
        print(f"Shape not found for removal: {shape_name}")
        return

    shape_element = shape._element
    shape_element.getparent().remove(shape_element)


def add_icon_in_placeholder(
    slide,
    placeholder_name,
    icon_name,
    width=0.90,
    remove_placeholder=True
):
    """
    Add and center an icon inside a named placeholder.

    The placeholder coordinates determine where the icon is placed.
    """

    placeholder = find_shape(slide, placeholder_name)

    if placeholder is None:
        print(f"Icon placeholder not found: {placeholder_name}")

        print("Available shapes:")

        for shape in slide.shapes:
            print(f"  - {shape.name}")

        return

    icon_path = BASE_DIR/"novaretail_icon_repo" / icon_name

    if not icon_path.is_file():
        print(f"Icon file not found: {icon_path}")
        return

    icon_width = Inches(width)
    icon_height = icon_width

    # Prevent the icon from becoming larger than the placeholder.
    maximum_width = int(placeholder.width * 0.85)
    maximum_height = int(placeholder.height * 0.85)

    if icon_width > maximum_width:
        icon_width = maximum_width

    if icon_height > maximum_height:
        icon_height = maximum_height

    # Center the icon horizontally and vertically.
    icon_left = (
        placeholder.left
        + (placeholder.width - icon_width) // 2
    )

    icon_top = (
        placeholder.top
        + (placeholder.height - icon_height) // 2
    )

    # Save the placeholder element before inserting the image.
    placeholder_element = placeholder._element

    slide.shapes.add_picture(
        str(icon_path),
        icon_left,
        icon_top,
        width=icon_width,
        height=icon_height
    )

    # Remove the dashed placeholder box.
    if remove_placeholder:
        placeholder_element.getparent().remove(
            placeholder_element
        )

    print(
        f"Icon added successfully: {icon_name} "
        f"inside {placeholder_name}"
    )
def set_comparison_chart(
    slide,
    shape_name,
    current_df,
    previous_df,
    metric,
    current_label,
    previous_label,
    title
):
    shape = find_shape(slide, shape_name)

    if shape is None:
        print(f"Chart placeholder not found: {shape_name}")
        return

    if metric not in current_df.columns:
        print(f"Metric not found in current quarter: {metric}")
        return

    if metric not in previous_df.columns:
        print(f"Metric not found in previous quarter: {metric}")
        return

    # Revenue by region for current quarter
    current_regional = (
        current_df
        .groupby("Region", as_index=False)[metric]
        .sum()
    )

    # Revenue by region for previous quarter
    previous_regional = (
        previous_df
        .groupby("Region", as_index=False)[metric]
        .sum()
    )

    # Combine both quarters
    comparison = pd.merge(
        previous_regional,
        current_regional,
        on="Region",
        how="outer",
        suffixes=("_previous", "_current")
    ).fillna(0)

    chart_data = CategoryChartData()

    chart_data.categories = (
        comparison["Region"]
        .astype(str)
        .tolist()
    )

    # Previous quarter
    chart_data.add_series(
        str(previous_label),
        comparison[f"{metric}_previous"]
        .astype(float)
        .tolist()
    )

    # Current quarter
    chart_data.add_series(
        str(current_label),
        comparison[f"{metric}_current"]
        .astype(float)
        .tolist()
    )

    # Save template chart position
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    # Remove template chart placeholder
    shape._element.getparent().remove(shape._element)

    # GROUPED / CLUSTERED chart
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left,
        top,
        width,
        height,
        chart_data
    ).chart

    chart.has_title = True
    chart.chart_title.text_frame.text = title

    # Required because we need to distinguish Q1 and Q2
    chart.has_legend = True
    chart.legend.include_in_layout = False

    chart.value_axis.has_major_gridlines = True

def create_regional_headline(
    current_df,
    previous_df,
    metric="Revenue_USD_M"
):
    current = (
        current_df
        .groupby("Region")[metric]
        .sum()
    )

    previous = (
        previous_df
        .groupby("Region")[metric]
        .sum()
    )

    comparison = pd.concat(
        [previous, current],
        axis=1,
        keys=["previous", "current"]
    ).dropna()

    if comparison.empty:
        return "Regional Revenue Performance"

    comparison["change"] = (
        comparison["current"]
        - comparison["previous"]
    )

    strongest = comparison["change"].idxmax()
    weakest = comparison["change"].idxmin()

    strongest_change = comparison.loc[
        strongest,
        "change"
    ]

    weakest_change = comparison.loc[
        weakest,
        "change"
    ]

    # If weakest region actually declined
    if weakest_change < 0:
        return (
            f"{strongest} Leads Revenue Growth "
            f"While {weakest} Declines"
        )

    # If every region grew
    return (
        f"{strongest} Leads Revenue Growth; "
        f"{weakest} Shows the Smallest Gain"
    )
def format_content_text(
    slide,
    shape_name,
    font_size,
    color=RGBColor(0, 0, 0),
    bold=False
):
    shape = find_shape(slide, shape_name)

    if shape is None or not shape.has_text_frame:
        print(f"Text shape not found for formatting: {shape_name}")
        return

    shape.text_frame.word_wrap = True

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.color.rgb = color
def create_customer_headline(
    current_df,
    previous_df
):
    current_nps = current_df["NPS"].mean()
    previous_nps = previous_df["NPS"].mean()

    change = current_nps - previous_nps

    if change > 0:
        return (
            f"Customer Advocacy Improves by "
            f"{abs(change):.1f} NPS Points"
        )

    elif change < 0:
        return (
            f"Customer Advocacy Declines by "
            f"{abs(change):.1f} NPS Points"
        )

    return "Customer Advocacy Remains Stable"
def create_risk_headline(
    current_df,
    previous_df
):
    current_stockout = (
        current_df["Stockout_Rate_pct"].mean()
    )

    previous_stockout = (
        previous_df["Stockout_Rate_pct"].mean()
    )

    change = current_stockout - previous_stockout

    if change > 0:
        return (
            f"Stockout Risk Increases by "
            f"{abs(change):.2f} Percentage Points"
        )

    elif change < 0:
        return (
            f"Stockout Risk Improves by "
            f"{abs(change):.2f} Percentage Points"
        )

    return "Stockout Risk Remains Stable"
    # -----------------------------------------------------
    # Slide 2: Regional Performance
    # -----------------------------------------------------

    slide2 = presentation.slides[1]

    regional_headline = create_regional_headline(
        p_now,
        p_prev,
        "Revenue_USD_M"
    )

    fill_shape(
        slide2,
        "title_2",
        regional_headline
    )

    # Standard title formatting
    format_content_text(
        slide2,
        "title_2",
        CONTENT_TITLE_SIZE,
        CONTENT_TITLE_COLOR,
        True
    )

    # Remove template chart instructions
    remove_shape(
        slide2,
        "chart_performance_label"
    )

    # Grouped comparison chart
    set_comparison_chart(
        slide2,
        "chart_performance",
        p_now,
        p_prev,
        "Revenue_USD_M",
        current_label=str(latest_q),
        previous_label=str(previous_q),
        title="Revenue by Region Across Quarters"
    )

    fill_shape(
        slide2,
        "insights_performance",
        f"• Revenue: {fmt(revenue, 'M')} USD in {latest_q}.\n"
        f"• Units sold: {fmt(units, 'K')}.\n"
        f"• Revenue QoQ change: {fmt(revenue_change, '%')}.\n"
        f"• Compare regional returns and marketing spend "
        f"with revenue performance."
    )

    # Standard insight formatting
    format_content_text(
        slide2,
        "insights_performance",
        CONTENT_INSIGHT_SIZE,
        CONTENT_TEXT_COLOR,
        False
    )

    add_icon_in_placeholder(
        slide2,
        "icon_2",
        "globe_region.png",
        width=CONTENT_ICON_SIZE,
        remove_placeholder=True
    )

    remove_shape(
        slide2,
        "icon_2_label"
    )   
def main():
    data,presentation=load_inputs()
    
    performance =data["Regional_Performance"]
    experience = data["Customer_Experience"]
    risk=data["Operational_Risk"]

    quarters = ordered_quarters(performance)
    if len(quarters)<2:
       raise ValueError("Atleast 2 quarters are required for Quarter over quarter analysis.")


    latest_q, previous_q= quarters[-1],quarters[-2]
    p_now = aggregate(performance,latest_q)
    p_prev= aggregate(performance,previous_q)
    c_now= aggregate(experience,latest_q)
    c_prev= aggregate(experience,previous_q)
    r_now = aggregate (risk,latest_q)
    r_prev = aggregate (risk,previous_q)

    revenue =p_now["Revenue_USD_M"].sum()
    previous_revenue = p_prev["Revenue_USD_M"].sum()
    units= p_now["Units_Sold_K"].sum()
    revenue_change = change_pct(revenue, previous_revenue)

    nps= c_now["NPS"].mean()
    csat= c_now["CSAT_pct"].mean()
    previous_nps =c_prev["NPS"].mean()
    nps_change=nps-previous_nps

    stockout =r_now["Stockout_Rate_pct"].mean()
    previous_stockout= r_prev["Stockout_Rate_pct"].mean()
    stockout_change=stockout- previous_stockout


    # -----------------------------------------------------
    # Slide 1: Executive Summary
    # -----------------------------------------------------

    slide1 = presentation.slides[0]

    fill_shape(
        slide1,
        "subtitle_1",
        f"{latest_q} Business Review"
    )

    fill_shape(
        slide1,
        "kpi1_value",
        fmt(revenue, "M")
    )

    fill_shape(
        slide1,
        "kpi1_label",
        "Revenue (USD)"
    )

    fill_shape(
        slide1,
        "kpi2_value",
        fmt(revenue_change, "%")
    )

    fill_shape(
        slide1,
        "kpi2_label",
        "Revenue QoQ Change"
    )

    fill_shape(
        slide1,
        "kpi3_value",
        fmt(nps)
    )

    fill_shape(
        slide1,
        "kpi3_label",
        "Average NPS"
    )

    fill_shape(
        slide1,
        "summary_text",
        f"In {latest_q}, revenue was "
        f"{revenue:.2f}M USD versus "
        f"{previous_revenue:.2f}M USD in "
        f"{previous_q}. Revenue changed by "
        f"{fmt(revenue_change, '%')}. "
        f"Average NPS was {fmt(nps)}, and the "
        f"stockout rate was {fmt(stockout, '%')}. "
        f"Review regional results and operational "
        f"indicators before finalizing business actions."
    )

    # Insert the icon into the correct template placeholder.
    add_icon_in_placeholder(
        slide1,
        "icon_1",
        "growth_chart.png",
        width=0.65,
        remove_placeholder=True
    )

    # Remove the template's icon instruction text.
    remove_shape(
        slide1,
        "icon_1_label"
    )

    # Apply white and bold formatting to Slide 1 text.
    slide1_white_shapes = [
        "subtitle_1",
        "kpi1_value",
        "kpi1_label",
        "kpi2_value",
        "kpi2_label",
        "kpi3_value",
        "kpi3_label",
        "summary_text",
    ]

    for shape_name in slide1_white_shapes:
        set_slide1_text_color(
        slide1,
        shape_name,
        color=RGBColor(255, 255, 255)
         )

    # -----------------------------------------------------
    # Slide 2: Regional Performance
    # -----------------------------------------------------
    

    slide2 = presentation.slides[1]

    regional_headline = create_regional_headline(
        p_now,
        p_prev,
        "Revenue_USD_M"
    )

    fill_shape(
        slide2,
        "title_2",
        regional_headline
    )

    # Standard title formatting
    format_content_text(
        slide2,
        "title_2",
        CONTENT_TITLE_SIZE,
        CONTENT_TITLE_COLOR,
        True
    )

    # Remove template chart instructions
    remove_shape(
        slide2,
        "chart_performance_label"
    )

    # Grouped comparison chart
    set_comparison_chart(
        slide2,
        "chart_performance",
        p_now,
        p_prev,
        "Revenue_USD_M",
        current_label=str(latest_q),
        previous_label=str(previous_q),
        title="Revenue by Region Across Quarters"
    )

    fill_shape(
        slide2,
        "insights_performance",
        f"• Revenue: {fmt(revenue, 'M')} USD in {latest_q}.\n"
        f"• Units sold: {fmt(units, 'K')}.\n"
        f"• Revenue QoQ change: {fmt(revenue_change, '%')}.\n"
        f"• Compare regional returns and marketing spend "
        f"with revenue performance."
    )

    # Standard insight formatting
    format_content_text(
        slide2,
        "insights_performance",
        CONTENT_INSIGHT_SIZE,
        CONTENT_TEXT_COLOR,
        False
    )

    add_icon_in_placeholder(
        slide2,
        "icon_2",
        "globe_region.png",
        width=CONTENT_ICON_SIZE,
        remove_placeholder=True
    )

    remove_shape(
        slide2,
        "icon_2_label"
    )

     
    # -----------------------------------------------------
    # Slide 3: Customer Experience
    # -----------------------------------------------------
    # Find highest and lowest NPS regions dynamically
    best_nps_row = c_now.loc[c_now["NPS"].idxmax()]
    worst_nps_row = c_now.loc[c_now["NPS"].idxmin()]

    best_region = best_nps_row["Region"]
    best_nps = best_nps_row["NPS"]

    worst_region = worst_nps_row["Region"]
    worst_nps = worst_nps_row["NPS"]

    # Insight-led headline
    fill_shape(
        presentation.slides[2],
        "title_3",
        f"Customer Advocacy Diverges Sharply, with {worst_region} "
        f"Trailing Other Regions"
    )

    # Chart
    set_chart(
        presentation.slides[2],
        "chart_customer",
        c_now,
        "NPS",
        f"NPS by Region | {latest_q}"
    )

    # Identify the other regions for supporting insight
    other_regions = c_now[
        ~c_now["Region"].isin([best_region, worst_region])
    ].sort_values("NPS", ascending=False)

    other_region_text = ", ".join(
        f"{row['Region']} ({row['NPS']:.0f})"
        for _, row in other_regions.iterrows()
    )

    # Dynamic insights
    fill_shape(
        presentation.slides[2],
        "insights_customer",

        f"• {best_region} leads customer advocacy with NPS "
        f"{best_nps:.0f}.\n"

        f"• {worst_region} trails significantly at NPS "
        f"{worst_nps:.0f}.\n"

        f"• Other regions remain stronger: "
        f"{other_region_text}.\n"

        f"• Overall average NPS improved from "
        f"{previous_nps:.2f} to {nps:.2f}, "
        f"a {nps_change:+.2f}-point QoQ change."
    )

       # -----------------------------------------------------
    # Slide 4: Operational Risk
    # -----------------------------------------------------

    slide4 = presentation.slides[3]

    risk_headline = create_risk_headline(
        r_now,
        r_prev
    )

    fill_shape(
        slide4,
        "title_4",
        risk_headline
    )

    # Same title formatting as Slides 2 and 3
    format_content_text(
        slide4,
        "title_4",
        CONTENT_TITLE_SIZE,
        CONTENT_TITLE_COLOR,
        True
    )

    remove_shape(
        slide4,
        "chart_risk_label"
    )

    set_chart(
        slide4,
        "chart_risk",
        r_now,
        "Stockout_Rate_pct",
        "Stockout Rate by Region"
    )

    fill_shape(
        slide4,
        "insights_risk",
        f"• Average stockout rate: {fmt(stockout, '%')}.\n"
        f"• Previous-quarter rate: "
        f"{fmt(previous_stockout, '%')}.\n"
        f"• Change: "
        f"{fmt(stockout_change, ' percentage points')}.\n"
        f"• Review return-processing time, warehouse "
        f"capacity utilization, and fraud incidents "
        f"for emerging risks."
    )

    # Same insight formatting
    format_content_text(
        slide4,
        "insights_risk",
        CONTENT_INSIGHT_SIZE,
        CONTENT_TEXT_COLOR,
        False
    )

    add_icon_in_placeholder(
        slide4,
        "icon_4",
        "risk_alert.png",
        width=CONTENT_ICON_SIZE,
        remove_placeholder=True
    )

    remove_shape(
        slide4,
        "icon_4_label"
    )
    # Slide 5: Recommended Actions
    # -----------------------------------------------------
    # Slide 5: Recommended Actions & Next Steps
    # -----------------------------------------------------

    slide5 = presentation.slides[4]

    # Main recommendations
    fill_shape(
        slide5,
        "recommendations_text",
        "1. Investigate regional revenue declines using actual sales "
        "and marketing data.\n\n"
        "2. Prioritize regions with declining NPS or CSAT and review "
        "delivery and support-ticket trends.\n\n"
        "3. Review stockout rates, warehouse utilization, and fraud "
        "incidents with operations owners.\n\n"
        "4. Assign accountable owners and monitor the same metrics "
        "next quarter."
    )

    # Add action icon
    add_icon_in_placeholder(
        slide5,
        "icon_5",
        "lightbulb_action.png",
        width=0.90,
        remove_placeholder=True
    )

    # Remove original icon instruction text
    remove_shape(
        slide5,
        "icon_5_label"
    )

    # Closing statement
    fill_shape(
        slide5,
        "closing_statement",
        "Use regional evidence to prioritize actions and validate "
        "their impact in the next quarterly review."
    )

    # Make recommendations white on dark-blue background
    set_slide1_text_color(
        slide5,
        "recommendations_text",
        color=RGBColor(255, 255, 255)
    )

    # Make closing statement white
    set_slide1_text_color(
        slide5,
        "closing_statement",
        color=RGBColor(255, 255, 255)
    )
    # Save the generated presentation

    presentation.save(OUTPUT_FILE)

    # Reopen the output to verify it is a valid PowerPoint

    verified = Presentation(OUTPUT_FILE)

    print("Deck generated successfully:", OUTPUT_FILE)
    print("Slides:", len(verified.slides))
    print("Quarter comparison:", previous_q, "to", latest_q)


if __name__ == "__main__":
    main()