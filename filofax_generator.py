from datetime import datetime, timedelta
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import math

# Set up page dimensions for Filofax sheet
FILOFAX_WIDTH = 12 * cm
FILOFAX_HEIGHT = 8 * cm
HOLE_PUNCH_MARGIN = 1 * cm

# Column dimensions
DAY_COL_WIDTH = 1 * cm
MONTH_COL_WIDTH = 3.5 * cm
WEEK_COL_WIDTH = 1.5 * cm
TOTAL_WIDTH = 7 * DAY_COL_WIDTH + MONTH_COL_WIDTH + WEEK_COL_WIDTH
MARGIN_LEFT = (FILOFAX_WIDTH - TOTAL_WIDTH) / 2

# Printing setup for Letter paper
LETTER_WIDTH, LETTER_HEIGHT = letter
CUT_MARGIN = 0.2 * cm  # 2mm cut margin
SHEETS_PER_PAGE = 3    # Three Filofax sheets per Letter page
SHEET_MARGIN = 1 * cm  # 1cm margin for printer
SHEET_SPACING = 0.2 * cm  # Space between sheets

def create_filofax_calendar(filename):
    # Create a PDF for printing on Letter paper
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Define date range (July 2025 to December 2026)
    start_date = datetime(2025, 7, 1)
    end_date = datetime(2026, 12, 31)
    
    # Find first Monday
    current_date = start_date
    while current_date.weekday() != 0:  # Monday is 0
        current_date += timedelta(days=1)
    
    # Track current year and sheet count
    current_year = current_date.year
    sheet_count = 0
    
    # Calculate total height for 3 sheets
    total_sheets_height = 3 * FILOFAX_HEIGHT + 2 * SHEET_SPACING
    top_margin = (LETTER_HEIGHT - total_sheets_height) / 2
    
    # Generate weekly pages
    while current_date <= end_date:
        # Check for new year
        if current_date.year != current_year:
            current_year = current_date.year
            sheet_count = 0  # Reset for new year
        
        # Calculate position on Letter paper
        page_position = sheet_count % SHEETS_PER_PAGE
        if page_position == 0:
            # Start a new Letter page after every 3 sheets
            c.showPage()
            # Reset sheet counter for new page
            sheet_count = 0
            page_position = 0
        
        # Calculate y position based on sheet position
        y_offset = LETTER_HEIGHT - top_margin - (page_position + 1) * FILOFAX_HEIGHT - page_position * SHEET_SPACING
        
        # Save state and translate to sheet position with right shift
        c.saveState()
        c.translate(SHEET_MARGIN, y_offset)
        
        # Draw the Filofax sheet
        draw_filofax_sheet(c, current_date)
        
        # Add cutting marks at the top
        draw_cutting_marks(c)
        
        # Restore state for next sheet
        c.restoreState()
        
        sheet_count += 1
        current_date += timedelta(days=7)
    
    # Save the final page
    c.save()

def draw_filofax_sheet(c, current_date):
    # Set up the Filofax sheet canvas
    c.setFont("Helvetica", 7)
    
    # Get week information
    week_num = current_date.isocalendar()[1]
    year = current_date.year
    
    # Get months in week
    months_in_week = set()
    for i in range(7):
        day_date = current_date + timedelta(days=i)
        months_in_week.add(day_date.month)
    
    # Create month string
    if len(months_in_week) == 1:
        month_str = (current_date + timedelta(days=3)).strftime('%B')  # Mid-week month
    else:
        # Get first and last month in week
        first_month = current_date.strftime('%b')
        last_month = (current_date + timedelta(days=6)).strftime('%b')
        month_str = f"{first_month} / {last_month}"
    
    # Draw day headers (Monday to Sunday)
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    for i, day in enumerate(days):
        x = MARGIN_LEFT + i * DAY_COL_WIDTH
        # Positioned 1mm below top line
        c.drawCentredString(x + DAY_COL_WIDTH/2, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 0.4*cm, day)
    
    # Draw dates
    for i in range(7):
        day_date = current_date + timedelta(days=i)
        x = MARGIN_LEFT + i * DAY_COL_WIDTH
        y = FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 0.9*cm
        
        # Draw date number
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + DAY_COL_WIDTH/2, y, str(day_date.day))
        c.setFont("Helvetica", 7)
    
    # Draw month in middle section
    month_x = MARGIN_LEFT + 7 * DAY_COL_WIDTH
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(month_x + MONTH_COL_WIDTH/2, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 0.5*cm, month_str)
    
    # Draw week number and year in right section
    week_x = month_x + MONTH_COL_WIDTH
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(week_x + WEEK_COL_WIDTH/2, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 0.5*cm, f"W{week_num}")
    c.setFont("Helvetica", 8)
    c.drawCentredString(week_x + WEEK_COL_WIDTH/2, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 0.9*cm, str(year))
    
    # Draw grid lines (extended to bottom)
    draw_grid(c)
    
    # Add bottom line for the sheet
    c.line(0, 0, FILOFAX_WIDTH, 0)
    
    # Add tiny horizontal lines every 0.5cm below title row
    draw_horizontal_lines(c)

def draw_grid(c):
    # Top horizontal line (hole punch margin)
    c.line(0, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN, FILOFAX_WIDTH, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)
    
    # Bottom of header row horizontal line
    c.line(0, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 1.0*cm, FILOFAX_WIDTH, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 1.0*cm)
    
    # Vertical lines (extended to bottom of page)
    # Left edge
    c.line(0, 0, 0, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)
    
    # Day columns
    for i in range(8):  # After each day (including start of first day)
        x = MARGIN_LEFT + i * DAY_COL_WIDTH
        c.line(x, 0, x, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)
    
    # After day columns
    month_x = MARGIN_LEFT + 7 * DAY_COL_WIDTH
    c.line(month_x, 0, month_x, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)
    
    # After month column
    week_x = month_x + MONTH_COL_WIDTH
    c.line(week_x, 0, week_x, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)
    
    # Right edge
    c.line(FILOFAX_WIDTH, 0, FILOFAX_WIDTH, FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN)

def draw_horizontal_lines(c):
    """Draw horizontal lines every 0.5cm below the title row"""
    # Start position (0.5cm below header row)
    start_y = FILOFAX_HEIGHT - HOLE_PUNCH_MARGIN - 1.0*cm - 0.5*cm
    num_lines = int((start_y) / (0.5 * cm))
    
    # Draw lines from top to bottom
    for i in range(num_lines):
        y_pos = start_y - (i * 0.5 * cm)
        if y_pos < 0:
            break
        c.line(0, y_pos, FILOFAX_WIDTH, y_pos)

def draw_cutting_marks(c):
    """Add cutting marks at the top of the sheet"""
    # Top left mark
    c.line(-0.2*cm, FILOFAX_HEIGHT, 0, FILOFAX_HEIGHT)  # Horizontal
    c.line(0, FILOFAX_HEIGHT, 0, FILOFAX_HEIGHT + 0.2*cm)  # Vertical
    
    # Top right mark
    c.line(FILOFAX_WIDTH, FILOFAX_HEIGHT, FILOFAX_WIDTH + 0.2*cm, FILOFAX_HEIGHT)  # Horizontal
    c.line(FILOFAX_WIDTH, FILOFAX_HEIGHT, FILOFAX_WIDTH, FILOFAX_HEIGHT + 0.2*cm)  # Vertical
    
    # Add bottom cutting marks too
    c.line(-0.2*cm, 0, 0, 0)  # Horizontal left
    c.line(0, 0, 0, -0.2*cm)  # Vertical left
    c.line(FILOFAX_WIDTH, 0, FILOFAX_WIDTH + 0.2*cm, 0)  # Horizontal right
    c.line(FILOFAX_WIDTH, 0, FILOFAX_WIDTH, -0.2*cm)  # Vertical right

if __name__ == "__main__":
    create_filofax_calendar("filofax_calendar_print_ready.pdf")
    print("Print-ready Filofax calendar generated successfully!")
