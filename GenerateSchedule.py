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

# Define the professor class
class Professor():

    def __init__(self):
        self.Id = 0
        self.FirstName = ''
        self.LastName = ''
        self.Availability = dict() # a dictionary mapping each time slot to a boolean, with True indicating the professor is available during that time slot, and False indicating that they are unavailable.

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

    # Instantiate the dictionary of professors
    Professors = dict()

    # Read the professor info into a data frame
    df = pd.read_excel(ProfessorInfoExcelFile)

    # Count the number of time slots
    NumTimeSlots = len(df.columns) - NonTimeColumns

    # Generate the list of time slots
    TimeSlots = range(NumTimeSlots)

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

        # Raise a warning
        print('Warning: No professor with the last name of %s could be found in the list of professors.' %ProfLastName)

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
                GetIdSuccess = True
            except ValueError:
                print('Warning: Could not retrieve the ID number corresponding to Professor %s appearing in the list of preferred faculty for visitor %s %s.' %(ProfLastName, v.FirstName, v.LastName))
                print('This entry in the list of preferred faculty will be ignored.  If you believe this is an error, please ensure that this professor appears in the Faculty Availability file.')
                GetIdSuccess = False

            # Check if the ID was successfully retrieved
            if GetIdSuccess == True:

                # Add the professor to the dictionary with the appropriate number of preference points
                v.PreferencePoints[ProfId] = MaxPreferencePoints - i

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
    model = pywraplp.Solver('simple_mip_program', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

    # Create the model variables
    print('\tDefining the decision variables...')
    Meeting = dict()
    for v in Visitors:
        for p in Professors:
            for t in TimeSlots:
                Meeting[(v,p,t)] = model.IntVar(0, 1, 'Visitor %d assigned to meet with Professor %d during time slot %d.' % (v, p, t))

    # Create the constraints
    print('\tDefining the constraints...')
    ## Each visitor can meet with at most one professor during any given time slot
    for v in Visitors:
        for t in TimeSlots:
            model.Add(
                sum(Meeting[(v,p,t)] for p in Professors) <= 1
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

    # Set the objective
    print('\tDefining the objective...')

    ## Define the weights of the various objectives
    Weight = {
        'Maximize the happiness points' : 1
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
    )

    # Return the model and the decision variable dictionary
    return (model, Meeting)

def PrintVisitorSchedule(Visitors, Professors, TimeSlots, Meeting, v):
    # Prints out the schedule for the specified visitor
    #
    # Inputs:
    #   v = the Id number of the visitor whose schedule you'd like to print out
    
    # Print out the visitor's name
    print('Visitor: %s %s' % (Visitors[v].FirstName, Visitors[v].LastName))

    # Loop over the time slots
    for t in TimeSlots:

        # Initialize the string to print
        PrintString = '\tPeriod %d:' % t

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

        # Print out the result
        print(PrintString)

def PrintAllVisitorSchedules(Visitors, Professors, TimeSlots, Meeting):

    # Loop over all the visitors
    for v in Visitors:

        # Print out the schedule for this visitor
        PrintVisitorSchedule(Visitors, Professors, TimeSlots, Meeting, v)

if __name__ == '__main__':

    # Import the visitor information
    Visitors = ImportVisitorInfo()

    # Import the professor and time slot information
    (Professors, TimeSlots) = ImportProfessorInfo()

    # Calculate the number of "preference points" that each visitor associates with each professor
    CalcPreferencePoints(Visitors, Professors)

    # Build the model
    (model, Meeting) = BuildModel(Visitors, Professors, TimeSlots)   

    # Solve the model
    print('Solving the model...')
    status = model.Solve()

    # Check for optimality
    if status == pywraplp.Solver.OPTIMAL:
        
        # Print a success message
        print('Success! The model was solved to optimality.')

    else:

        # Display an error message
        print('Error: The problem does not have an optimal solution.')

    # Print out all the visitors' schedules
    PrintAllVisitorSchedules(Visitors, Professors, TimeSlots, Meeting)
