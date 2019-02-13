
	Simplified Add-In Creation 
	==========================

	This Visual Studio template creates a new Inventor add-in.  It's different from the 
	standard add-in template that is delivered with Inventor and installed as part of the 
	developer tools. It adds some additional functionality to simplify adding commands.
	
	To use it, follow the steps below.
	
	1) Copy the VBNiftyAddInTemplateVS2017.zip file to:
	
	    C:\Users\<username>\Documents\Visual Studio 2017\Templates\ProjectTemplates\Visual Basic

	2) Create a new project, selecting the template.
	
	3) Build the project.
	   Building the project will also copy the necessary files to:
	   
	   %appdata%\Autodesk\ApplicationPlugins
	   
	   Inventor looks in this folder for add-ins.
	   
	4) Start Inventor. You'll get an "Add-In Manager Security Alert" dialog notifying
       	   you that the add-in couldn't be authenticated and has been blocked.
	   Open the Add-In Manager, select the add-in from the list and uncheck the "Block" 
	   check box. Check the "Loaded/Unloaded" and "Load Automatically" check boxes and 
	   click "OK".
	
	5) Create a new part or open any existing part document. Select the "Tools" tab. 
	   Look for the "Sample" tab in the ribbon and the "Command Name" command. Running 
	   the command will display a dialog showing the version of Inventor you're running.
	   
	To remove the add-in, delete the add-in folder from location specified in step 3 above.