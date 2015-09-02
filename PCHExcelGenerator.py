
"""
	author: Kristofer Ken Castro
	date: 9/2/2015

	This class generates an excel spreadsheet with 4 columns:
	ID, EMPLOYEE, MANAGER, SALES.  
	Three dimensions and one measure. The dataset created by this class 
	contains a Parent-child hiearchy between the employees and their managers.

	The algorithm uses two CSV files to generate full names: 
	a ~5000 line file containg first names and a ~80000 line file containing last names.
"""

import xlsxwriter
from Employee import Employee
import csv
import random
import sys

class PCHExcelGenerator:

	#constants
	NUM_ROOT_NODES = 1;
	NUM_ROOT_NODES_WITH_RANDOM_LEVELS = 1;
	MIN_LEVEL = 1;
	MAX_LEVEL = 1;
	MIN_CHILD_PER_NODE = 1;
	MAX_CHILD_PER_NODE = 1;

	workbook = xlsxwriter.Workbook('OrgHier.xlsx');
	worksheet = workbook.add_worksheet();

	# used to keep track of our row and column in the excel table
	row = 0;
	col = 0;

	#table of first names
	firstNameDatabase = [];
	lastNameDatabase = [];

	rootEmployees = [];

	def __init__(self, numRootNodes, numRootNodesWithRandomLevels, minLevel, maxLevel, minChildPerNode, maxChildPerNode):
		self.NUM_ROOT_NODES = numRootNodes;
		self.NUM_ROOT_NODES_WITH_RANDOM_LEVELS = numRootNodesWithRandomLevels;
		self.MIN_LEVEL = minLevel;
		self.MAX_LEVEL = maxLevel;
		self.MAX_CHILD_PER_NODE = maxChildPerNode;
		self.MIN_CHILD_PER_NODE = minChildPerNode;
		self.createFirstRow();
		self.loadFirstNameDatabase();
		self.loadLastNameDatabase();

	def execute(self):
		self.createEmployees();
		self.cleanUp();

	def cleanUp(self):
		self.workbook.close();

	def createFirstRow(self):
		self.worksheet.write(self.row, self.col, 'ID');
		self.worksheet.write(self.row, self.col+1, 'EMPLOYEE');
		self.worksheet.write(self.row, self.col+2, 'MANAGER');
		self.worksheet.write(self.row, self.col+3, 'SALES');
		self.row+=1;

	def createEmployees(self):
		#create root employees
		for x in range (0, self.NUM_ROOT_NODES):
			newEmployee = self.createRandomEmployee();
			self.rootEmployees.append(newEmployee);
		selectedEmployees = self.createSelectedEmployees();
		for employee in selectedEmployees:
			self.createLevelsForEmployee(employee);

	def createRandomEmployee(self, manager=None):
		firstName = self.generateRandomFirstName();
		lastName = self.generateRandomLastName();
		newEmployee = Employee(firstName, lastName, manager);
		self.createEmployeesToManage(newEmployee);
		self.writeToExcel(newEmployee);

		return newEmployee;

	def createRandomEmployeeToManage(self, manager):
		firstName = self.generateRandomFirstName();
		lastName = self.generateRandomLastName();
		newEmployee = Employee(firstName, lastName, manager);
		self.writeToExcel(newEmployee);
		return newEmployee;

	#create root employees that we will generate random amount of levels in
	def createSelectedEmployees(self):
		numOfRootEmployees = len(self.rootEmployees);
		numToSelect = self.NUM_ROOT_NODES_WITH_RANDOM_LEVELS;
		allEmployeesIndex = [];
		allSelectedEmployees = [];
		if (numOfRootEmployees == numToSelect): return self.rootEmployees;

		# first make an array to keep track of index when we move things around
		for x in range (0, numOfRootEmployees):
			allEmployeesIndex.append(x);

		# randomly select employees from the poop and make sure
		# we don't consider the selected ones again.
		# from tail+1 to end of the list is our list of selected employees
		tail = numOfRootEmployees-1;
		for x in range (0, numToSelect):
			randSelect = random.randrange(0, tail+1);
			# perform a swap
			temp = allEmployeesIndex[tail];
			allEmployeesIndex[tail] = allEmployeesIndex[randSelect];
			allEmployeesIndex[randSelect] = tail;
			# add the selected employee to our list
			allSelectedEmployees.append(self.rootEmployees[allEmployeesIndex[tail]]);
			tail-=1;

		return allSelectedEmployees;

	def createLevelsForEmployee(self, employee):
		numOfLevels = random.randrange(self.MIN_LEVEL, self.MAX_LEVEL+1);
		if (numOfLevels == 0 ): return; # do nothing since level 0 is just the node itself
		lastManager = employee;

		nextLevelEmployee = self.createRandomEmployee(lastManager);
		lastManager.addEmployeeToManage(nextLevelEmployee);
		lastManager = nextLevelEmployee;
		if (numOfLevels == 1 ): return; # we done

		for x in range (0, numOfLevels-1):
			nextLevelEmployee = self.createRandomEmployee(lastManager);
			lastManager.addEmployeeToManage(nextLevelEmployee);
			lastManager = nextLevelEmployee;

	# create a random number of employee the given employee has to manage
	def createEmployeesToManage(self, employee):
		randNumChild = random.randrange(self.MIN_CHILD_PER_NODE, self.MAX_CHILD_PER_NODE+1);
		numChildAlready = len(employee.employeesManaging);
		for x in range (numChildAlready, randNumChild):
			newEmployeeToManage = self.createRandomEmployeeToManage(employee);
			employee.addEmployeeToManage(newEmployeeToManage);

	def generateRandomFirstName(self):
		length = len(self.firstNameDatabase);
		return self.firstNameDatabase[random.randrange(0,length)];

	def generateRandomLastName(self):
		length = len(self.lastNameDatabase);
		return self.lastNameDatabase[random.randrange(0,length)];

	def loadFirstNameDatabase(self):
		with open("CSV_Database_of_First_Names.csv", newline='') as csvfile:
			spamreader = csv.reader(csvfile, delimiter=' ');
			for row in spamreader:
				self.firstNameDatabase.append(row[0]);


	def loadLastNameDatabase(self):
		with open("CSV_Database_of_Last_Names.csv", newline='') as csvfile:
			spamreader = csv.reader(csvfile, delimiter=' ');
			for row in spamreader:
				self.lastNameDatabase.append(row[0]);

	# write the employee data to excel
	def writeToExcel(self, employee):
		self.worksheet.write(self.row, self.col, employee.id);
		self.worksheet.write(self.row, self.col+1, employee.getFullName());
		self.worksheet.write(self.row, self.col+2, (employee.manager.id if (employee.manager != None) else ""));
		self.worksheet.write(self.row, self.col+3, employee.sales);
		self.row+=1





# prompt for paramters

numRootNodes = input("Enter # of root nodes: ");
while not numRootNodes.isdigit(): numRootNodes = input("\tEnter again.  Must be a digit: ");

numRootNodesWithRandomLevels = input("Enter # of root nodes with random # of levels: ");
while not numRootNodesWithRandomLevels.isdigit() or numRootNodesWithRandomLevels > numRootNodes:
	numRootNodesWithRandomLevels = input("Enter again.  Must be a digit and less than {0}:".format(numRootNodes));

minLevel = input("Enter the minimum # of levels for the random root node: ");
while not minLevel.isdigit() : minLevel = input("\tEnter again.  Must be a digit: ");

maxLevel = input("Enter the maximum # of levels for the random root node: ");
while not maxLevel.isdigit() : maxLevel = input("\tEnter again.  Must be a digit: ");

minChild = input("Enter the minimum # of child per node: ");
while not minChild.isdigit() : minChild = input("\tEnter again.  Must be a digit: ");

maxChild = input("Enter the maximum # of child per node: ");
while not maxChild.isdigit() : maxChild = input("\tEnter again.  Must be a digit: ");

#run the generator
generator = PCHExcelGenerator(int(numRootNodes), int(numRootNodesWithRandomLevels), int(minLevel), int(maxLevel),int(minChild),int(maxChild));
generator.execute();
print("finished generating excel!");


