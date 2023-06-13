import os
from skimage import io
import pytesseract
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pypdf import PdfMerger

# anticipated dpi of scanned image
dpi = 300

# scan horizontal text based on y-position as bottom of text
#-------------------------------------------------------------------------
def scan_text(image, y):
    y1 = y-50
    y2 = y+20
    ocr_text = pytesseract.image_to_string(image[y1:y2, (dpi*1):(dpi*7.5)])
    return ocr_text

# scan horizontal bubbles based on y-position as bottom of bubbles and text
#-------------------------------------------------------------------------
def scan_bubbles(image, y):
    
    r = dpi*0.1    # radius of bubble
    low = 255      # lowest average intensity of bubble
    selection = 0  # most filled selection (1-4)

    # iterate through possible bubbles
    for i in range(4):
        x = dpi*(0.75*i+4)
        avg = np.mean(image[(y-r/2):(y+3*r/2), (x-r/2):(x+r/2)])

        if (avg < low) and avg < 150:
            selection = i+1
            low = avg

    return selection





















# Get the name of the project with the most filled in bubble
#-------------------------------------------------------------------------
def get_project(image):
    projects = {'Project A':[565, 435],
                'Project B':[565, 535],
                'Project C':[565, 635],
                'Project D':[565, 735],
                'Project E':[565, 835],
                'Project F':[565, 935]}

    best_score = 255
    name = ''

    for project in projects:

        score = get_score(image, projects[project][0], projects[project][1])
        if score < min(200, best_score):
            name = project
            best_score = score

    return name

# Get the name of the room with the most filled in bubble
#-------------------------------------------------------------------------
def get_room(image):
    rooms = {'Room A':[1545, 435],
             'Room B':[1545, 535],
             'Room C':[1545, 635],
             'Room D':[1545, 735]}

    best_score = 255
    name = ''

    for room in rooms:
        score = get_score(image, rooms[room][0], rooms[room][1])

        if score < min(200, best_score):
            name = room
            best_score = score

    return name

# Get a table of the scores for each category with the most filled in bubbles
#-------------------------------------------------------------------------
def get_ranks(image):
    cols = {'1 - Needs Improvement':943,
            '2 - Emerging':1533,
            '3 - Exceptional':1907}

    rows = {'Problem Statement':1200, 
            'Slide Quality and Clarity':1300,
            'Organization':1400,
            'Clinical Assessment':1500,
            'Scientific Assessment':1600,
            'Market Assessment':1700,
            'Risk Assessment':1800,
            'Analysis':1900,
            'Citation of Sources':2000}
    
    ranks = {'Problem Statement':'', 
             'Slide Quality and Clarity':'',
             'Organization':'',
             'Clinical Assessment':'',
             'Scientific Assessment':'',
             'Market Assessment':'',
             'Risk Assessment':'',
             'Analysis':'',
             'Citation of Sources':''}
        
    for row in rows:

        best_score = 255
        rank = ''

        for col in cols:

            score = get_score(image, cols[col], rows[row])

            if score < min(200, best_score):
                rank = col
                best_score = score

        ranks[row] = rank

    return ranks

# Get comments
#-------------------------------------------------------------------------
def get_comments(image):
    return image[2200:2740, 320:2265]


#-------------------------------------------------------------------------
# Read and return the results for all the scans in a directory
#-------------------------------------------------------------------------

ranks = {'Problem Statement':'', 
         'Slide Quality and Clarity':'',
         'Organization':'',
         'Clinical Assessment':'',
         'Scientific Assessment':'',
         'Market Assessment':'',
         'Risk Assessment':'',
         'Analysis':'',
         'Citation of Sources':''}

results = pd.DataFrame(columns = ['Team'] + list(ranks))

# location of scanned images
path = os.path.dirname(__file__)

i = 1
for filename in os.listdir(path):
    if filename.endswith(".jpeg") or filename.endswith(".jpg"):

        # import scanned image
        scan = io.imread(path + '/' + filename)

        # identify team informormation
        current_project = get_project(scan)
        current_room = get_room(scan)
        current_team = current_room + '-' + current_project

        # get team scores
        current_result = get_ranks(scan)
        current_result['Team'] = current_team
        pd.Series(current_result)

        # append to all scores
        results = results._append(current_result, ignore_index=True)

        # save iamge of comments
        io.imsave(path + '/' + current_team + ' - text ' + str(i) + '.pdf', get_comments(scan))
        i += 1

#Save all results to excel file
results.to_excel(path + '/' + 'rubric results.xlsx')


#-------------------------------------------------------------------------
# Analyze and plot results
#-------------------------------------------------------------------------

teams = list(results['Team'].unique())
for team in teams:
    # filter results by team
    team_results =  results[results['Team'] == team]

    # get value counts for each team
    result_counts = team_results[list(ranks)].apply(pd.value_counts)
    result_counts = result_counts.fillna(0)
    result_counts = result_counts.transpose()

    # plot value counts and save
    result_counts.plot(kind='barh', stacked=True)
    plt.title(team)
    plt.tight_layout()
    plt.savefig(path + '/' + team + ' - scores.pdf')

#-------------------------------------------------------------------------
# Combine into PDF
#-------------------------------------------------------------------------

for team in teams:

    merger = PdfMerger()
    
    for filename in os.listdir(path):
        print(filename)
        if team in filename:
            merger.append(path + '/' + filename)

    merger.write(path + '/' + team +'.pdf')
    merger.close()