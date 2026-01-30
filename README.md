Step 1: Create the list of Users

let
    //First, enter data using the "Enter Data" function. Enter the full names of the users.
    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45W8ixOTExWcM5ILMpJLVaK1YlWCsnPVXAqSkypBPO88oszFBxzclLzoNxUBafSoqL8cjA3ILGkKDM5W8E3MSM/F6Q/FgA=", 
      BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [User = _t]),

    #"Changed Type" = Table.TransformColumnTypes(Source, {{"User", type text}}),

    //When working with people with multiple names split the name apart.
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Split", each Text.Split([User], " ")),

    //Create additional rows for each part of the name.
    #"Expanded Split" = Table.ExpandListColumn(#"Added Custom", "Split"),

    //Create a simple matching boolean statement to determine if the end of the split word matches the end of the full name. This will help us identify the correct last name.
    #"Added Custom1" = Table.AddColumn(#"Expanded Split", "Match", each if Text.End([User], 1) = Text.End([Split], 1) then 1 else 0),

    //Filter the rows for matching ending letters.
    #"Filtered Rows" = Table.SelectRows(#"Added Custom1", each ([Match] = 1)),

    //Create the style of user name that best matches your file path. As an example, here is a created file path for our team in our fake company.
    #"Create usernames" = Table.AddColumn(#"Filtered Rows", "Users", each [Split] & Text.Start([User], 1) ),
    //Remove all columns but the Users column
    #"Removed Columns" = Table.RemoveColumns(#"Create usernames", {"Split", "Match","User"}),

    //Lastly, turn the table users into a list. There is where understanding of OOP comes in handy to know that lists are really powerful tools to call objects and use zero-baseed counting.
    Name = #"Removed Columns"[Users]

in

    Name
Step 2: Create the User Parameter
In Power Query, navigate to Home > Manager Parameters > New Parameter

Name the parameter- using similar logic suggested in OOP, name this user to denote that a single “user” from the list “users” can be entered one at a time. Feel free to add a description. 
Next, select Type = Text, Suggested Value = Query, and select your list Users. Lastly, enter a current value which can be your username name to test the functionality of the parameter. 
Select OK and to return to the query editor window. 
Creating “User” Parameter:
"BradyT" meta [IsParameterQuery=true, ExpressionIdentifier=Users, Type="Text", IsParameterQueryRequired=true]

Step 3: Build your query before your source
This can be the most complicated step in the process but incredibly rewarding when done correctly. Now it’s time to build your connection string and pull in some data. 
Our variable BasePath will be a concatenated string including the full file path folder where our data lives with a forward slash after the folder name. 
Between this where the username exists, we will replace it with the parameter User. Below is an example of what it should look like. 
The source variable will then be made up of our variable BasePath and the name of the data source we are accessing. 
As mentioned this has been tested for CSV, Excel, and folder sources, but I can imagine lots of functionality across multiple sources saved in shared drives.

let
BasePath = "Redacted:\forprivacy\" & User & "\to\preventhacking\",
    Source = Excel.Workbook(File.Contents(BasePath & "Dataset.xlsx"), null, true),

Lastly, conduct all your normal transformations as standard and test with a peer to ensure functionality. This will boost productivity and allow for seamless workflow across teams when working on pbix files in shared drive locations while connecting to data sources. 
