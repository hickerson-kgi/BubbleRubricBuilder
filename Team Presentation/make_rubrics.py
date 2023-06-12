import os
import pandas as pd
from reportlab.pdfgen import canvas
# ReportLab https://docs.reportlab.com/reportlab/userguide/ch1_intro/

#-------------------------------------------------------------------------
# Load and Organize Rubric Criteria from Excel
#-------------------------------------------------------------------------
path = os.path.dirname(__file__)
raw_data = pd.read_excel(path+'/rubric_criteria.xlsx')

team_criteria = list(raw_data['Team Criteria'].dropna())
team_ratings = list(raw_data['Team Ratings'].dropna())
individual_criteria = list(raw_data['Individual Criteria'].dropna())
individual_ratings = list(raw_data['Individual Ratings'].dropna())

#-------------------------------------------------------------------------
# Load and Organize Team Membership from Excel
#-------------------------------------------------------------------------

# read excel file
path = os.path.dirname(__file__)
raw_data = pd.read_excel(path+'/team_membership.xlsx')

# organize by teams into dictionary
team_membership = {}
list_teams = list(raw_data['Team'].unique())

for team in list_teams:
    members = raw_data[raw_data['Team'] == team]
    members = list(members['Student'])
    team_membership[team] = members

#-------------------------------------------------------------------------
# Generate Rubric for PDF
# https://docs.reportlab.com/reportlab/userguide/ch2_graphics/
#-------------------------------------------------------------------------
path = os.path.dirname(__file__)
c = canvas.Canvas(filename = path+'/presentation_rubrics.pdf',
                  bottomup = 0,
                  pagesize=(612, 792))

# Return page pixel given inch dimensions and 72px/inch
#-------------------------------------------------------------------------
def inch(inches):
    return int(72*inches)

# Draw a horizontal line at vertical position in inches
#-------------------------------------------------------------------------
def hline(v_position):
    c.line(inch(1), inch(v_position), inch(7.5), inch(v_position))

# draw a horizontal row for a criteria and ratings
#-------------------------------------------------------------------------
def draw_row(criteria, ratings, v_position):
    c.drawString(inch(1), inch(v_position), criteria)

    i = 1
    for rating in ratings:
        r = inch(0.1)
        c.circle(    inch(0.75*i+4),    inch(v_position)-r, r, stroke=1, fill=0)
        c.drawString(inch(0.75*i+4.25), inch(v_position), rating[0])
        i += 1

# Draw Team Information
#-------------------------------------------------------------------------

for team in team_membership:
    # Current vertical position on page in inches
    v_current = 1

    # Team Title
    c.drawString(inch(1),inch(v_current),'Team Score: ' + team)
    hline(v_current+ 0.125)

    # Team Scores
    v_current += 0.5
    for criteria in team_criteria:
        draw_row(criteria, team_ratings, v_current)
        v_current += 0.25

    # Team Comments
    c.rect(inch(1), inch(v_current), inch(6.5), inch(4), stroke=1, fill=0) 
    v_current += 0.25
    c.drawString(inch(1.125),inch(v_current),'Team Comments')

    
    # Draw Individual Information
    #-------------------------------------------------------------------------
    # New Page
    c.showPage()
    v_current = 1
    
    for member in team_membership[team]:

        if v_current > 9.5:
            c.showPage()
            v_current = 1

        # Individual Name
        v_current += 0.25
        c.drawString(inch(1),inch(v_current),"Individual Score: "+
                    team + ' - ' + member)
        hline(v_current+ 0.125)

        # Individual Scores
        v_current += 0.5
        for criteria in individual_criteria:
            draw_row(criteria, individual_ratings, v_current)
            v_current += 0.25

        # Individual Comments
        c.rect(inch(1), inch(v_current), inch(6.5), inch(1.4), stroke=1, fill=0) 
        v_current += 0.25
        c.drawString(inch(1.125),inch(v_current), member + ' Comments')
        v_current += 1.4
        
    # Finish Current Page
    c.showPage() 

# Save File
c.save()