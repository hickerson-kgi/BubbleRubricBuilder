import os
from skimage import io
import pytesseract
import numpy as np

import pandas as pd
import matplotlib.pyplot as plt
from pypdf import PdfMerger

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# Return page pixel given inch dimensions and 300dpi
#-------------------------------------------------------------------------
def in2px(inches):
    return int(300*inches)

# scan horizontal text based on y-position as bottom of text
#-------------------------------------------------------------------------
def scan_text(image, y):
    y1 = y-60
    y2 = y+30
    ocr_text = pytesseract.image_to_string( image[y1:y2, in2px(1):in2px(7.5)] )
    return ocr_text

# scan horizontal bubbles based on y-position as bottom of bubbles and text
#-------------------------------------------------------------------------
def scan_bubbles(image, y):
    
    r = in2px(0.1)    # radius of bubble
    low = 255      # lowest average intensity of bubble
    selection = 0  # most filled selection (1-4)

    # iterate through possible bubbles
    for i in range(4):
        x = in2px(0.5*i+4.75)
        avg = np.mean(image[int(y-3*r/2):int(y-r/2), int(x-r/2):int(x+r/2)])

        if (avg < low) and avg < 200:
            selection = i+1
            low = avg
    
    return selection

# Get comments
#-------------------------------------------------------------------------
def scan_team_comments(image):
    return image[in2px(3.75):in2px(7.75), in2px(1):in2px(7.5)]

# Get team data
#-------------------------------------------------------------------------
def get_team_results(image, team_criteria, team_ratings):

    # initialize team results
    team_results = {}
    
    # determine team name
    team_name = scan_text(image, in2px(1))
    team_name = team_name[:-1]
    team_results['Team'] = team_name

    # determine criteria scores
    y = in2px(1.5)
    for criteria in team_criteria:
        selection = scan_bubbles(image, y)

        if selection:
            team_results[criteria] = team_ratings[selection-1]

        y += in2px(0.25)

    return pd.Series(team_results)

# Get individual data
#-------------------------------------------------------------------------
def get_individual_results(image, individual_criteria, individual_ratings):
    individual_results = {}


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
# Read and return the results for all the scans in a directory
#-------------------------------------------------------------------------

# location of scanned images
path = os.path.dirname(__file__)
filenames = ['scan_ahickers_2023-06-13-12-14-58_1.jpeg',
             'scan_ahickers_2023-06-13-12-14-58_3.jpeg']

results = pd.DataFrame()

i = 1
for filename in filenames:
    if filename.endswith(".jpeg") or filename.endswith(".jpg"):
        # import scanned image
        scan = io.imread(path + '/' + filename)

        # Get team data
        current_team_results = get_team_results(scan, team_criteria, team_ratings)
        results = results._append(current_team_results, ignore_index=True)

        # save iamge of comments
        io.imsave(path + '/' + current_team_results['Team'] + ' - comment_' + str(i) + '.pdf', scan_team_comments(scan))
        i += 1

results.to_csv(path + '/' + 'rubric results.csv')


#-------------------------------------------------------------------------
# Analyze and plot results
#-------------------------------------------------------------------------

teams = list(results['Team'].unique())
for team in teams:
    # filter results by team
    team_results =  results[results['Team'] == team]

    # get value counts for each team
    result_counts = team_results[team_criteria].apply(pd.value_counts)
    result_counts = result_counts.fillna(0)
    result_counts = result_counts.transpose()

    # plot value counts and save
    result_counts.plot(kind='barh', stacked=True)
    plt.title(team)
    plt.tight_layout()
    plt.savefig(path + '/' + team + ' - barplot.pdf')

#-------------------------------------------------------------------------
# Combine into PDF
#-------------------------------------------------------------------------

for team in teams:

    merger = PdfMerger()
    
    for filename in os.listdir(path):
        
        if team in filename:
            merger.append(path + '/' + filename)

    merger.write(path + '/' + team +'.pdf')
    merger.close()
