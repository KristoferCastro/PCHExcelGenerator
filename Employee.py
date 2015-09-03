import random
import uuid;

class Employee:
	id = "";
	firstName = "";
	lastName = "";
	sales = 0;
	manager = None;
	employeesManaging = [];

	def __init__(self, firstName, lastName, manager):
		self.firstName = firstName;
		self.lastName = lastName;
		self.sales = self.generateSalesNumber();
		self.manager = manager;
		self.id = self.generateId();
		self.employeesManaging = [];

	def getFullName(self):
		return self.firstName + " " + self.lastName;

	def generateSalesNumber(self):
		return random.randrange(0, 99);

	def setMinManageEmployee(self, min):
		self.minManageEmployee = min;

	def setMaxManageEmployee(self, max):
		self.maxManageEmployee = max;

	def generateId(self):
		return self.firstName + "." + self.lastName + "-" + str(uuid.uuid4())[:8];

	def addEmployeeToManage(self, employee):
		self.employeesManaging.append(employee);

	def __repr__(self):
		return self.getFullName();
