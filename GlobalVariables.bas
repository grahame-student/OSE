Attribute VB_Name = "GlobalVariables"
Option Explicit
DefObj A-Z

' Global variables to be used through out the application
Global FF As Integer ' File handle to use for reads and writes of save files
Global SaveFileData As ESS ' The data structure instance that holds the loaded data

