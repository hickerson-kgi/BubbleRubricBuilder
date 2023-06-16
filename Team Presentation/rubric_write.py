import os
import pandas as pd
from reportlab.pdfgen import canvas
# ReportLab https://docs.reportlab.com/reportlab/userguide/ch1_intro/

#-------------------------------------------------------------------------
# Load and Organize Rubric Criteria from CSV file
#-------------------------------------------------------------------------
path = os.path.dirname(__file__)
raw_data = pd.read_csv(path+'/rubric_criteria.csv')

team_criteria = list(raw_data['Team Criteria'].dropna())
team_ratings = list(raw_data['Team Ratings'].dropna())
individual_criteria = list(raw_data['Individual Criteria'].dropna())
individual_ratings = list(raw_data['Individual Ratings'].dropna())

#-------------------------------------------------------------------------
# Load and Organize Team Membership from CSV file
#-------------------------------------------------------------------------

# read excel file
path = os.path.dirname(__file__)
raw_data = pd.read_csv(path+'/team_membership.csv')

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
                  bottomup = 0,         # top left is (0,0) coordinate
                  pagesize=(612, 792))  # 8.5 x 11 in sheet at 72 dpi

# Return page pixel given inch dimensions and 72px/inch
#-------------------------------------------------------------------------
def in2px(inches):
    return int(72*inches)

# Draw a horizontal line at vertical position in inches
#-------------------------------------------------------------------------
def hline(v_position):
    c.line(in2px(1), in2px(v_position), in2px(7.5), in2px(v_position))

# draw a horizontal row for a criteria and ratings
#-------------------------------------------------------------------------
def draw_row(criteria, ratings, v_position):
    c.drawString(in2px(1), in2px(v_position), criteria)

    i = 0
    for rating in ratings:
        r = in2px(0.1)
        c.circle(    in2px(0.5*i+4.75),    in2px(v_position)-r, r, stroke=1, fill=0)
        c.drawString(in2px(0.5*i+4.9), in2px(v_position), str(i+1))
        i += 1

# Draw Team Information
#-------------------------------------------------------------------------

for team in team_membership:
    # Current vertical position on page in inches
    v_current = 1
    
    # Team Title
    c.drawString(in2px(1),in2px(v_current), team)
    print('Team Title', v_current)
    hline(v_current+ 0.125)

    # Team Scores
    v_current += 0.5
    for criteria in team_criteria:
        draw_row(criteria, team_ratings, v_current)
        print('Team Criteria', v_current)
        v_current += 0.25

    # Team Comments
    c.rect(in2px(1), in2px(v_current), in2px(6.5), in2px(4), stroke=1, fill=0) 
    print('Team Comments', v_current)
    v_current += 0.25
    c.drawString(in2px(1.125),in2px(v_current),'Team Comments')

    
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
        c.drawString(in2px(1),in2px(v_current), team + ' - ' + member)
        print('Name', v_current)
        hline(v_current+ 0.125)

        # Individual Scores
        v_current += 0.5
        for criteria in individual_criteria:
            draw_row(criteria, individual_ratings, v_current)
            print('Individual Criteria', v_current)
            v_current += 0.25

        # Individual Comments
        c.rect(in2px(1), in2px(v_current), in2px(6.5), in2px(1.4), stroke=1, fill=0) 
        print('Individual Comments', v_current)
        v_current += 0.25
        c.drawString(in2px(1.125),in2px(v_current), member + ' Comments')
        v_current += 1.4
        
    # Finish Current Page
    c.showPage() 

# Save File
c.save()