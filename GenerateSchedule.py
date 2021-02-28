# Import the OR-Tools library
from ortools.linear_solver import pywraplp

# Import Pandas
import pandas as pd
from pandas import ExcelFile

# Import the os module
import os

# Define the visitor class
class Visitor():

    def __init__(self):
        self.Id = 0
        self.FirstName = ''
        self.LastName = ''
        self.PreferredProfessors = []  # the rank-ordered list of preferred professors, with the first professor in the list being the most preferred.
        self.PreferencePoints = dict() # a dictionary which maps professor id numbers to preference points.  The more desirable the professor, the greater the points.  If a professor is absent from this dictionary, they are assumed to be associated with zero preference points.
        self.Availability = '' # This indicates when the visitor is available for meetings.  Three values are valid: 'morning', 'afternoon', and 'na'.
        self.Happiness = 0
        self.NumberOfMeetings = 0

# Define the professor class
class Professor():

    def __init__(self):
        self.Id = 0
        self.FirstName = ''
        self.LastName = ''
        self.Availability = dict() # a dictionary mapping each time slot to a boolean, with True indicating the professor is available during that time slot, and False indicating that they are unavailable.
        self.NumberOfMeetingsAvailable = 0

# Define the function for reading in the visitor information
def ImportVisitorInfo():

    # Specify the name of the excel file
    VisitorInfoExcelFile = 'Visitor Preferences.xlsx'

    # Print out a status update
    print('Attempting to import the visitor information from \"%s%s%s\"...' %(os.getcwd(), os.sep, VisitorInfoExcelFile))

    # Instantiate the dictionary of visitors
    Visitors = dict()

    # Read the visitor preferences into a data frame
    df = pd.read_excel(VisitorInfoExcelFile)

    # Loop over all the rows of the data frame
    for (i, row) in df.iterrows():

        # Instantiate a new visitor
        v = Visitor()

        # Specify the index for this visitor
        v.Id = i

        # Retrieve the visitor's first and last name
        v.FirstName = row['First']
        v.LastName = row['Last']

        # Retrieve the list of preferred professors as a string
        PreferredProfessorsString = row['Faculty list']

        # Retrieve the visitor's availability
        v.Availability = str(row['Category'])
        
        # Convert the string into a list
        v.PreferredProfessors = PreferredProfessorsString.split(', ')

        # Add this visitor to the growing dictionary of visitors
        Visitors[v.Id] = v

    # Print out a status update
    print('Successfully read in the information of %d visitors.' %len(Visitors) )

    # Return the result
    return Visitors

# Define the function for importing the professor information
def ImportProfessorInfo():

     # Specify the name of the excel file
    ProfessorInfoExcelFile = 'Faculty Availability.xlsx'

    # Specify the number of columns that do not correspond to a time slot
    NonTimeColumns = 1

    # Print out a status update
    print('Attempting to import the professor information from \"%s%s%s\"...' %(os.getcwd(), os.sep, ProfessorInfoExcelFile))

    # Read the professor info into a data frame
    df = pd.read_excel(ProfessorInfoExcelFile)

    # Count the number of time slots
    NumTimeSlots = len(df.columns) - NonTimeColumns

    # Generate the dictionary of time slots
    TimeSlots = dict()
    for t in range(NumTimeSlots):

        # Add an entry for the current time slot
        TimeSlots[t] = str(df.columns[NonTimeColumns + t])

    # Instantiate the dictionary of professors
    Professors = dict()

    # Loop over all the rows of the data frame
    for (i, row) in df.iterrows():

        # Instantiate a new professor
        p = Professor()

        # Specify the index for this visitor
        p.Id = i

        # Retrieve the professor's first and last name
        NameString = row['Faculty']
        NameList = NameString.split(', ')
        p.FirstName = NameList[1]
        p.LastName = NameList[0]

        # Loop over the time slots
        for t in TimeSlots:

            # Extract the professor's availability for the current time slot
            if row[NonTimeColumns + t] == 1:
                p.Availability[t] = True
            else:
                p.Availability[t] = False

        # Add this professor to the growing dictionary of professors
        Professors[p.Id] = p

    # Print out a status update
    print('Successfully read in the information of %d professors.' %len(Professors) )

    # Return the result
    return (Professors, TimeSlots)

# Define the function for looking up a professor's ID number
def GetProfID(ProfLastName):
    # Returns the ID number associated with the given prof's last name

    # Generate the list of professor last names
    ProfLastNames = [p.LastName for p in Professors.values()]

    # Check if this last name is in the list.
    if ProfLastName in ProfLastNames: # it's there

        # Find the position of the last name in the list of professor last names
        ListIndex = ProfLastNames.index(ProfLastName)

        # Find the ID number corresponding to this list index
        ProfId = [p.Id for p in Professors.values()][ListIndex]

    else: # it's not there

        # Raise an error
        raise ValueError

        # Specify a clearly incorrect value for the ID number
        ProfId = -1

    # Return the result
    return ProfId

# Define the function for calculating the number of "preference points" that each visitor associates with each professor
def CalcPreferencePoints(Visitors, Professors):
    # Builds up the dictionary of preference points for each visitor.

    # Specify the maximum number of preference points
    MaxPreferencePoints = 10

    # Instantiate the list of unrecognized professors
    UnrecognizedProfs = []

    # Loop over each of the visitors
    for v in Visitors.values():

        # Initialize all dictionary of preference points with all zeros
        for p in Professors:
            v.PreferencePoints[p] = 0

        # Loop over each of the professors in this visitor's list of preferred professors.
        for i in range(len(v.PreferredProfessors)):

            # Extract the professor corresponding to the current index
            ProfLastName = v.PreferredProfessors[i]

            # Lookup the id number corresponding to this professor
            try:
                ProfId = GetProfID(ProfLastName)
                
            except ValueError:

                # Set the flag to indicate failure in the attempt to get the ID
                GetIdSuccess = False

                # check if the prof's last name has already been added to the list of unrecognized professors
                if not ProfLastName in UnrecognizedProfs:

                    # Print a warning message
                    print('Warning: The name \"%s\" was found in the list of preferred professors for visitor %s %s and perhaps others.  However, no availability information was found for this professor.' %(ProfLastName, v.FirstName, v.LastName))
                    print('         This entry in the list of preferred professors will be ignored.  If you believe this is an error, please ensure that this professor appears in the Faculty Availability file.')
                
                    # Add the prof to the list
                    UnrecognizedProfs.append(ProfLastName)

            else:

                # Raise the flag to indicate success in getting the ID
                GetIdSuccess = True

            # Check if the ID was successfully retrieved
            if GetIdSuccess == True:

                # Add the professor to the dictionary with the appropriate number of preference points
                v.PreferencePoints[ProfId] = 1

        # Print the results for the current visitor
        #PrintPreferencePoints(v, Professors)

def PrintPreferencePoints(Visitor, Professors):
    # Input:
    #   Visitor = the object of the visitor whose preference points you'd like to print
    #   Professors = the dictionary of professors

    # Print out the visitor's name
    print(Visitor.FirstName + ' ' + Visitor.LastName)

    # Loop over each entry of their preference point dictionary
    for p in Visitor.PreferencePoints:

        # Initialize the print string with the professor's name
        PrintString = '\t' + Professors[p].LastName

        # Add the preference points
        PrintString += ': %d' % Visitor.PreferencePoints[p]

        # Print out the result
        print(PrintString)

# Define the function for building the optimization model
def BuildModel(Visitors, Professors, TimeSlots):
    # This function builds the constraint programming model for the problem
    # Inputs:
    #   Visitors = a dictionary of visitors.
    #   Professors = a dictionary of professors
    #   TimeSlots = a list of time slot indices
    # Outputs:
    #   model = a CP model object populated with decision variables, constraints, and an objective.

    # Print a status update
    print('Defining the optimization model...')

    # Instantiate the model
    model = pywraplp.Solver.CreateSolver('cbc')

    # Create the model variables
    print('\tDefining the decision variables...')

    ## Primary decision variables
    Meeting = dict()
    for v in Visitors:
        for p in Professors:
            for t in TimeSlots:
                Meeting[(v,p,t)] = model.IntVar(0, 1, 'Visitor %d assigned to meet with Professor %d during time slot %d.' % (v, p, t))

    ## Secondary decision variables
    MinMeetings = model.NumVar(0, model.infinity(), 'Minimum Meetings per Visitor')

    MinHappiness = model.NumVar(0, model.infinity(), 'Minimum Happiness per Visitor')

    # Create the constraints
    print('\tDefining the constraints...')
    ## Each visitor can meet with at most one professor during any given time slot
    for v in Visitors:
        for t in TimeSlots:
            model.Add(
                sum(Meeting[(v,p,t)] for p in Professors) <= 1
            )

    ## Visitors can only attend meetings permitted by their timezones
    for v in Visitors:
        if Visitors[v].Availability == 'morning':
            for t in TimeSlots:
                if t >= 8:
                    for p in Professors:
                        model.Add(
                            Meeting[v,p,t] == 0
                        )

        elif Visitors[v].Availability == 'afternoon':
            for t in TimeSlots:
                if t < 8:
                    for p in Professors:
                        model.Add(
                            Meeting[v,p,t] == 0
                        )

    ## Each professor can meet with at most one visitor during any given time slot
    for p in Professors:
        for t in TimeSlots:
            model.Add(
                sum(Meeting[(v,p,t)] for v in Visitors) <= 1
            )

    ## Each professor can only meet with visitors when the professor is available
    for p in Professors:
        for t in TimeSlots:
            for v in Visitors:
                model.Add(
                    Meeting[(v,p,t)] <= Professors[p].Availability[t]
                )

    ## Each professor-visitor pair can meet at most once
    for p in Professors:
        for v in Visitors:
            model.Add(
                sum(Meeting[(v,p,t)] for t in TimeSlots) <= 1
            )

    ## Each visitor must have at least the minimum number of meetings
    for v in Visitors:
        model.Add(
            sum(
                sum(
                    Meeting[(v,p,t)] 
                    for t in TimeSlots
                )
                for p in Professors
            )
            >=
            MinMeetings
        )

    ## Each visitor must have at least the minimum happiness score
    for v in Visitors:
        model.Add(
            sum(
                sum(
                    Meeting[(v,p,t)] * Visitors[v].PreferencePoints[p]
                    for t in TimeSlots
                )
                for p in Professors
            )
            >=
            MinHappiness
        )

    ## Each visitor must have at least the minimum number of free periods
    RequiredFreePeriods = 8
    for v in Visitors:
        model.Add(
            sum(
                sum(
                    Meeting[(v,p,t)] 
                    for t in TimeSlots
                )
                for p in Professors
            )
            <=
            len(TimeSlots) - RequiredFreePeriods
        )

    # Set the objective
    print('\tDefining the objective...')

    ## Define the weights of the various objectives
    Weight = {
        'Maximize the happiness points' : 1,
        'Maximize the number of meetings' : 0.1,
        'Maximize the minimum number of meetings': 1,
        'Maximize the minimum happiness score': 1
    }

    ## Define the objective
    model.Maximize(

        # Maximize the happiness points
        Weight['Maximize the happiness points'] *
        sum(
            sum(
                sum(
                    Meeting[(v,p,t)] * Visitors[v].PreferencePoints[p]
                    for t in TimeSlots
                )
                for v in Visitors
            )
            for p in Professors
        )

        +

        Weight['Maximize the number of meetings'] *
        sum(
            sum(
                sum(
                    Meeting[(v,p,t)]
                    for t in TimeSlots
                )
                for v in Visitors
            )
            for p in Professors
        )

        +

        Weight['Maximize the minimum number of meetings'] * MinMeetings

        +

        Weight['Maximize the minimum happiness score'] * MinHappiness
    )

    # Return the model and the decision variable dictionary
    return (model, Meeting)

def PrintVisitorSchedule(Visitors, Professors, TimeSlots, Meeting, v):
    # Prints out the schedule for the specified visitor
    #
    # Inputs:
    #   v = the Id number of the visitor whose schedule you'd like to print out
    
    # Build the name of the output file
    FileName = 'Visitor %s %s\'s Schedule.txt' % (Visitors[v].FirstName, Visitors[v].LastName)

    # Build the path to the directory
    FileDirectory = '%s%s%s' %(os.getcwd(), os.sep, 'Visitor Schedules')

    # Check if the directory exists
    if os.path.isdir(FileDirectory) == False:

        # Create the directory
        os.mkdir(FileDirectory)

    # Build the path to the output file
    FilePath = '%s%s%s' %(FileDirectory, os.sep, FileName)

    # Open up the file for writing
    File = open(FilePath, 'w')

    # Print out the visitor's name
    File.write('Visitor: %s %s\n' % (Visitors[v].FirstName, Visitors[v].LastName))

    # Loop over the time slots
    for t in TimeSlots:

        # Initialize the string to print
        PrintString = '\t%s (Period %d):' % (TimeSlots[t], t)

        # Initialize a flag to indicate that a meeting has not yet been found
        MeetingFound = False

        # Loop over the Professors
        for p in Professors:

            # check if the current visitor has a meeting scheduled with the current professor
            if Meeting[(v,p,t)].solution_value() == 1: # then a meeting between this visitor and professor has been scheduled

                # Add the professor's name to the print string
                PrintString += ' Professor %s' % Professors[p].LastName

                # Raise the flag to indicate that a meeting was found
                MeetingFound = True

        # Check if a meeting was found
        if MeetingFound == False: # then no meeting was found

            # Extend the print string to indicate free time
            PrintString += ' Free time'

        # Add a newline character
        PrintString += '\n'

        # Print out the result
        File.write(PrintString)

    # Close the file you were writing to
    File.close()

def PrintAllVisitorSchedules(Visitors, Professors, TimeSlots, Meeting):

    # Loop over all the visitors
    for v in Visitors:

        # Print out the schedule for this visitor
        PrintVisitorSchedule(Visitors, Professors, TimeSlots, Meeting, v)

def PrintProfessorSchedule(Visitors, Professors, TimeSlots, Meeting, p):
    # Prints out the schedule for the specified visitor
    #
    # Inputs:
    #   p = the Id number of the professor whose schedule you'd like to print out
    
     # Build the name of the output file
    FileName = 'Professor %s\'s Schedule.txt' % Professors[p].LastName

    # Build the path to the directory
    FileDirectory = '%s%s%s' %(os.getcwd(), os.sep, 'Professor Schedules')

    # Check if the directory exists
    if os.path.isdir(FileDirectory) == False:

        # Create the directory
        os.mkdir(FileDirectory)

    # Build the path to the output file
    FilePath = '%s%s%s' %(FileDirectory, os.sep, FileName)

    # Open up the file for writing
    File = open(FilePath, 'w')

    # Print out the professor's name
    File.write('Professor: %s\n' % Professors[p].LastName)

    # Loop over the time slots
    for t in TimeSlots:

        # Initialize the string to print
        PrintString = '\t%s (Period %d):' % (TimeSlots[t], t)

        # Initialize a flag to indicate that a meeting has not yet been found
        MeetingFound = False

        # Loop over the visitors
        for v in Visitors:

            # check if the current visitor has a meeting scheduled with the current professor
            if Meeting[(v,p,t)].solution_value() == 1: # then a meeting between this visitor and professor has been scheduled

                # Add the visitor's name to the print string
                PrintString += ' Visitor %s %s' % (Visitors[v].FirstName, Visitors[v].LastName)

                # Raise the flag to indicate that a meeting was found
                MeetingFound = True

        # Check if a meeting was found
        if MeetingFound == False: # then no meeting was found

            # Check if the professor is available during this time slot
            if Professors[p].Availability[t] == True:

                # Extend the print string to indicate free time
                PrintString += ' Free time (available)'

            else:

                # Extend the print string to indicate unavailability
                PrintString += ' Unavailable'

        # Add a newline character
        PrintString += '\n'

        # Print out the result
        File.write(PrintString)

    # Close the file
    File.close()

def PrintAllProfessorSchedules(Visitors, Professors, TimeSlots, Meeting):

    # Loop over all the professors
    for p in Professors:

        # Print out the schedule for this professor
        PrintProfessorSchedule(Visitors, Professors, TimeSlots, Meeting, p)

def CalcVisitorHappiness(Visitors, Professors, TimeSlots, Meeting):

    # Calculate the happiness of each visitor
    for v in Visitors:

        # Loop over their list of professors
        for p in Professors:

            # Check if they were assigned a meeting with that professor
            if sum(Meeting[(v,p,t)].solution_value() for t in TimeSlots) == 1:  # They were assigned a meeting with that professor

                # Increment their happiness accordingly
                Visitors[v].Happiness += Visitors[v].PreferencePoints[p]

                # Increment their meeting count accordingly
                Visitors[v].NumberOfMeetings += 1

def CalcMeetingsAvailable(Professors, TimeSlots):

    # Loop over the list of professors
    for p in Professors:

        # Loop over the time slots
        for t in TimeSlots:

            # Check if the prof is available during this time slot
            if Professors[p].Availability[t] == True:  # they are available

                # Increment their count of meetings available
                Professors[p].NumberOfMeetingsAvailable += 1

def PrintSummaryStatistics(Visitors, Professors, TimeSlots, Meeting):
    # This function prints some statistics to help assess the quality of the meeting assignments

    # Get the list of happiness scores
    HappinessScores = [Visitors[v].Happiness for v in Visitors]

    # Import the statistics module
    import statistics as stats

    # Print the mean happiness
    print('The mean happiness score is: %f' % stats.mean(HappinessScores))

    # Print the median happiness
    print('The median happiness score is: %f' % stats.median(HappinessScores))

    # Print the max happiness
    print('The max happiness score is: %f' % max(HappinessScores))

    # Print the min happiness
    print('The min happiness score is: %f' % min(HappinessScores))

    # Find the least happy visitors
    for v in Visitors:

        # Check if they are among the least happy
        if Visitors[v].Happiness == min(HappinessScores):

            # Print the visitor's name
            print('\tCheck %s %s' % (Visitors[v].FirstName, Visitors[v].LastName))

    # Print the standard deviation happiness
    print('The standard deviation in happiness scores is: %f' % stats.stdev(HappinessScores))

    # Get the list of meeting counts
    MeetingCounts = [Visitors[v].NumberOfMeetings for v in Visitors]

    # Print the mean number of meetings
    print('The mean number of meetings is: %f' % stats.mean(MeetingCounts))

    # Print the median happiness
    print('The median number of meetings is: %f' % stats.median(MeetingCounts))

    # Print the max happiness
    print('The max number of meetings is: %f' % max(MeetingCounts))

    # Print the min happiness
    print('The min number of meetings is: %f' % min(MeetingCounts))

    # Find the visitors with the least meetings
    for v in Visitors:

        # Check if they are among those with the fewest meetings
        if Visitors[v].NumberOfMeetings == min(MeetingCounts):

            # Print the visitor's name
            print('\tCheck %s %s' % (Visitors[v].FirstName, Visitors[v].LastName))

    # Print the standard deviation in the number of meetings
    print('The standard deviation in the number of meetings is: %f' % stats.stdev(MeetingCounts))

    # Get the list of number of meetings available for each prof
    MeetingsAvailable = [Professors[p].NumberOfMeetingsAvailable for p in Professors]

    # Calculate the total number of meetings available
    TotalMeetingsAvailable = sum(MeetingsAvailable)
    print('Total meetings with faculty available: %d' % TotalMeetingsAvailable)

    # Calculate the total number of meetings arranged
    TotalMeetingsScheduled = sum(MeetingCounts)
    print('Total meetings with faculty scheduled: %d' % TotalMeetingsScheduled)

    # Calculate the fraction of available meetings scheduled
    print('Percentage of available meetings scheduled: %.1f%%' % (float(TotalMeetingsScheduled)/float(TotalMeetingsAvailable)*100))

    # Calculate the minimum number of meetings that you could guarantee each student
    import math
    print('Given the faculty availability, the minimum number of meetings we could guarantee each visitor is: %d' % math.floor(TotalMeetingsAvailable / len(Visitors)))
        

if __name__ == '__main__':

    # Import the visitor information
    Visitors = ImportVisitorInfo()

    # Import the professor and time slot information
    (Professors, TimeSlots) = ImportProfessorInfo()

    # Calculate the number of "preference points" that each visitor associates with each professor
    CalcPreferencePoints(Visitors, Professors)

    # Build the model
    (model, Meeting) = BuildModel(Visitors, Professors, TimeSlots)   

    # Enable output
    model.EnableOutput()

    # Set the time limit
    MaxMinutes = 1
    model.set_time_limit(round(1000*60*MaxMinutes))

    # Solve the model
    print('Solving the model...')
    status = model.Solve()

    # Check for optimality
    if status == pywraplp.Solver.OPTIMAL:

        # Display a success message
        print('Optimal solution found!')

    elif status == pywraplp.Solver.INFEASIBLE:

        # Display an error message
        print('Error: The model was found to be infeasible.')
        exit()

    elif status == pywraplp.Solver.NOT_SOLVED:

        # Display an error message
        print('Error: The model was not solved to completion.  Consider increasing the amount of time allowed to solve the model.')
        exit()

    elif status == pywraplp.Solver.UNBOUNDED:

        # Display an error message
        print('Error: The model was found to be unbounded.')
        exit()

    elif status == pywraplp.Solver.ABNORMAL:

        # Display an error message
        print('Error: The solver exited with an abnormal status. Consider increasing the amount of time allowed to solve the model.')
        exit()

    else:

        # Give a status update
        print('Warning: The model was not solved to completion.  Calculating the optimality gap...')

        # Calculate the optimality gap
        OptimalityGap = model.Objective().BestBound() - model.Objective().Value()
        RelativeOptimalityGap = OptimalityGap / max(model.Objective().BestBound(), 0.001)

        # Print the optimality gap
        print('The optimality gap is %f%%' % (RelativeOptimalityGap * 100))

        if RelativeOptimalityGap > 0.01:

            # Display an error message
            print('Error: I was unable to solve the model to the desired precision in the time allotted. Consider increasing the amount of time allowed to solve the model.')
            exit()

        else:

            # Print a partial success message
            print('The model was solved to within an acceptable optimality gap.')

    # Print a success message
    print('Success!')

    # Calculate visitor happiness
    CalcVisitorHappiness(Visitors, Professors, TimeSlots, Meeting)

    # Count the number of meetings each prof is available
    CalcMeetingsAvailable(Professors, TimeSlots)

    # Print out some summary statistics
    PrintSummaryStatistics(Visitors, Professors, TimeSlots, Meeting)
    
    # Print out all the visitors' schedules
    print('Writing out the schedule for each visitor...')
    PrintAllVisitorSchedules(Visitors, Professors, TimeSlots, Meeting)

    # Print out all the professors' schedules
    print('Writing out the schedule for each professor...')
    PrintAllProfessorSchedules(Visitors, Professors, TimeSlots, Meeting)

    # Print a final message
    print('All done! Please inspect the individual schedules that were created in the \"Visitor Schedules\" and \"Professor Schedules\" directories.')
