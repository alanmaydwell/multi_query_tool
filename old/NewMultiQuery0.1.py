import Tkinter
import ttk
import tkMessageBox
import pyodbc,tkFileDialog,os,time
import pickle


"""
GUI for setting up details (username, database, password, SQL ) for
multiple SQL queries

NewQuery v0.1 - not finished. Has working loading and saving of query set parameters.
NewQuery v0.2 - Rows can be added and removed
        Filename for group of queries can be set. Details saved, details loaded.
        Constructs connection string and extracts SQL when run button pressed
NewQuery v0.3 - Added parameter field panel and handling. Copes with 0 to 6 fields.
Also has heading for each which is included in saved/loaded data.
Doesn't yet actually run any queries.

[New name]
NewMultiQuery v0.1 - Now runs queries.
Pasted in Querydata class. Made some changes to it - mainly due to way Tkinter imported here
Now can run queries and export results. Includes reading of parameters.
Added "safety catch" for quickload/quickslave
Should be fully working version.

[Improvements to make -
hange export format to plain text?
]
"""


class QueryPanel:
    """Tkinter panel holding SQL query details.
    Holds varibale numebr of querys (description, database, username,password, sql file0
    Supports multiple parameters
    Can load/save setup to disc.


    Args:
        title - Heading for frame
        defaultfile - default filename for query set

    """
    def __init__(self,frame,title="SQL Query Definitions",defaultfile='default.dat'):

        style = ttk.Style()
        style.configure("BW.TLabel",foreground="black",background="white", relief="groove")

        self.frame=frame

        self.title = ttk.Label(self.frame,text=title, font =("Arial",9,"bold") )
        self.title.grid(sticky="W")

        #Set load/save buttons
        self.setfile=os.path.join(os.getcwd(),defaultfile)#defualt filename

        fileframe=ttk.Frame(self.frame, relief="groove",borderwidth=6)
        fileframe.grid(columnspan=10,sticky="W")

        filelabel=ttk.Label(fileframe,text="Load / Save", font =("Arial",8, "bold"))
        filelabel.grid(row=0,column=0,sticky="W",pady=0,padx=0)

        self.bfilename=ttk.Button(fileframe,text="Load Query Set",command=self.set_filename)
        self.bfilename.grid(row=1,column=0,padx=2)

        self.bfilename=ttk.Button(fileframe,text="Save Query Set",command= lambda: self.set_filename("save"))
        self.bfilename.grid(row=1,column=1,padx=2)


        self.bload=ttk.Button(fileframe,text="Quick Load Query Set",command=self.loadfields)
        self.bload.grid(row=1,column=3,padx=2)

        self.bsave=ttk.Button(fileframe,text="Quick Save Query Set",command=self.savefields)
        self.bsave.grid(row=1,column=4,padx=2)

        #"Safety catch" tickbox for quickload and quicksave

        self.safelabel=ttk.Label(fileframe,text="Enable Quick Load/Save")
        self.safelabel.grid(row=0,column=3)
        self.safevar=Tkinter.IntVar()
        self.safevar.set(0)
        self.safebox=ttk.Checkbutton(fileframe, variable = self.safevar, command = self.safetycatch)
        self.safebox.grid(row=0,column=4,sticky="W")
        self.safetycatch()#

        #Show current filename
        self.flabel=ttk.Label(fileframe,text="Selected File:")
        self.flabel.grid(row=1,column=5,padx=4,stick="E")
        self.lfilename=ttk.Label(fileframe,text=os.path.basename(self.setfile),style="BW.TLabel")
        self.lfilename.grid(row=1,column=6,sticky="W")


        #Parameter fields
        row = self.frame.grid_size()[1]+1
        self.paramframe=ttk.Frame(self.frame,relief="groove",borderwidth=6)
        self.paramframe.grid(row=row,columnspan=10,sticky="W")
        paramlabel=ttk.Label(self.paramframe,text="Parameter Fields", font =("Arial",8, "bold"))
        paramlabel.grid(row=0,column=0,sticky="W",pady=0,padx=0)

        #Number of parameters drop-down list
        #Note can't associate callback with list itself but can with the associated variable
        self.param_entries=[]#list to hold entry fields.
        self.param_entry_names=[]#list to hold title text
        self.paramvar=Tkinter.StringVar()
        self.paramvar.trace("w", self.paramupdate)

        param=ttk.Combobox(self.paramframe,textvariable=self.paramvar, state="readonly", width=3)
        param.grid(row=0,column=1,sticky="S")
        param["values"]=("0","1","2","3","4","5","6")
        param.current(0)

        toprowlabel=ttk.Label(self.paramframe,text="Title(s)")
        toprowlabel.grid(row=0,column=2)
        botrowlabel=ttk.Label(self.paramframe,text="Value(s)")
        botrowlabel.grid(row=1,column=2)

        #Queries
        #SETUP QFRAME

        row=self.frame.grid_size()[1]+1
        self.qframe=ttk.Frame(self.frame,relief="groove",borderwidth=6)
        self.qframe.grid(row=row,columnspan=10,sticky="W")

        #Title
        row=self.qframe.grid_size()[1]+1

        qtitle=ttk.Label(self.qframe,text="Query Setup", font =("Arial",8, "bold"))
        qtitle.grid(row=row,column=0)

        #"Add" button
        self.bnew=ttk.Button(self.qframe,text="Add Row",command=self.add_row)
        self.bnew.grid(row=row,column=1,padx=4)

        #Column headings for query rows
        self.headings = ["Include","Description","Database","Username","Password","SQL Query"]
        row=self.qframe.grid_size()[1]+1
        for i,h in enumerate(self.headings):
            l=ttk.Label(self.qframe,text=h, font =("Arial",8,"bold"))
            l.grid(row=row,column=i)

        #Stores important data
        self.qr=[]#query row widget values (and SQL content)
        self.qrd=[]#extracted query row data from entry fields (not including SQl content)


    def add_query(self):
        """Adds row of query fields to GUI

        """

        details={}
        details["id"]=len(self.qr)#id for row. Want to be same as index of position in self.qr
        details["sql"]=""#will hold loaded SQL text (not the filename)


        #Add items benieth each heading
        row=self.qframe.grid_size()[1]+1

        col=0
        for h in (self.headings):
            #Tickbox
            if h=="Include":
                details["tickvar"]=Tkinter.IntVar()
                details["tickvar"].set(0)
                details["tick"]=ttk.Checkbutton(self.qframe, variable = details["tickvar"])
                details["tick"].grid(row=row,column=col)

            #SQL
            elif h=="SQL Query":
                details["qbutton"] = ttk.Button(self.qframe,text="Load", command = lambda id=details["id"]:self.sql_choose(id))
                details["qbutton"].grid(row=row, column=col, padx=3)
                col=col+1
                details[h]=ttk.Entry(self.qframe,width=48)
                details[h].grid(row=row,column=col)

            #Default Entry Fields
            else:
                details[h]=ttk.Entry(self.qframe)
                details[h].grid(row=row,column=col)

            col=col+1

        #Add delete button
        details["delbutton"] = ttk.Button(self.qframe,text="Delete Row", command = lambda id=details["id"]:self.delete_row(id))
        details["delbutton"].grid(row=row,column=col,padx=3)

        self.qr.append(details)

        #Make sure row ids are right (might not always be necessary)
        self.update_row_ids()
        #self.status = ttk.Label(self.qframe,text="12345678",style="BW.TLabel",padding=1)
        #self.status.grid(row=0,column=3)


    def add_row(self):
        """Adds new query row
        Currently simply calls self.add_query()
        """
        self.add_query()

    def update_row_ids(self):
        """Update "id" of each self.qr dictionary to match its position in self.qr list
        Also updates the callback command for each delete and query button

        "id" is important for each row's delete button
        """

        for i, q in enumerate(self.qr):
            q["id"]=i
            #Update reference in delbutton
            q["delbutton"].config(command = lambda id=i:self.delete_row(id))
            #Update referece in load button
            q["qbutton"].config(command = lambda id=i:self.sql_choose(id))


    def delete_row(self,id):
        """Deletes row. Used by Delete buttons in GUI

        Args:
            id - id value of self.qrd["id"] value of row to be deleted
        """

        #Don't delete last row
        if len(self.qr)<2:
            return

        #delete widgets from row
        for e in self.qr[id]:
            if  self.qr[id][e].__class__ in [ttk.Entry,ttk.Button,ttk.Label,ttk.Checkbutton]:
                self.qr[id][e].destroy()
        #Delete
        del(self.qr[id])

        #Not bothering to delete any self.qrd values for row
        #as self.qrd values are always created afresh

        #Update row id values following deletion
        self.update_row_ids()


    def readfields(self):
        """Reads values from fields for all rows
        stores values in self.qrd (query row data)

        Also reads SQL files if filenames set
        """
        self.qrd=[] # clear old query row data

        for q in self.qr:
            data={}
            for element in q:

                #Read values from Entry Fields and tick-box variables
                if q[element].__class__ in [ttk.Entry,Tkinter.IntVar]:
                    data[element]=q[element].get()

            self.qrd.append(data)


    def set_filename(self,mode=None):
        """Sets the filename for the query set.
        Loads or saves the file depending on mode.
        Filename set using Tkinter dialogue.

        Args:
            mode - when "save" save, otherwise load
        """
        if mode=="save":
            #Get filename when saving file
            filepath=tkFileDialog.asksaveasfilename(filetypes=[("DAT",".dat"),("All",".*")]
                ,title="Save - filename for Query Set",defaultextension=".dat")

        else:
            #Get "Load" filename
            filepath=tkFileDialog.askopenfilename(filetypes=[("DAT",".dat"),("All",".*")]
                ,title="Load - filename for Query Set")

        if filepath:
            #Set the filename value
            self.setfile=filepath
            #Update displayed name
            self.lfilename["text"]=os.path.basename(filepath)

            #Save file
            if mode=="save":
                self.savefields()
            #Default to load file
            else:
                #load the file
                self.loadfields()


    def safetycatch(self):
        """Enables/disables Quicksave and Quickload buttons
        based on whether tickbox is ticked"""
        if self.safevar.get():
            self.bload['state']='normal'
            self.bsave['state']='normal'
        else:
            self.bload['state']='disabled'
            self.bsave['state']='disabled'


    def savefields(self):
        """Save fields"""
        self.readfields()
        #Save configuration using pickle
        with open(self.setfile,"wb") as wf:
            #Save row details and number of parameter fields, parameter entry names
            descriptions=[e.get() for e in self.param_entry_names]
            values=[e.get() for e in self.param_entries]
            pickle.dump([self.qrd,self.paramvar.get(),descriptions,values],wf)


    def loadfields(self):
        """Loads values to entry fields"""

        #Read file
        with open(self.setfile,"rb") as rf:
            self.qrd, np, headings, values = pickle.load(rf)

        #Set number of parameters fields
        self.paramvar.set(np)

        #Restore parameter valuess
        for e, text in zip(self.param_entries,values):
            e.delete(0,Tkinter.END)
            e.insert(0,text)

        #Restore parameter headings
        for e, text  in zip(self.param_entry_names,headings):
            e.delete(0,Tkinter.END)
            e.insert(0,text)

        #Add extra rows if there aren't enough
        shortage=len(self.qrd)-len(self.qr)
        if shortage>0:
            for i in range(0,shortage):
                self.add_query()
        #Remove surplus rows
        if shortage<0:
            for i in range(0,-1*shortage):
                self.delete_row(id=-1+len(self.qr))#this way to always delete last row

        #Populate fields with loaded values
        for i,row in enumerate(self.qrd):
            for q in row:
                element=self.qr[i][q]#correspondence between qrd and qr relies on key names for entry fields matching
                #Entry fieldss
                if element.__class__ == ttk.Entry:
                    element.delete(0,Tkinter.END)
                    element.insert(0,row[q])
                #Variables assoicated with tick-boxes
                elif element.__class__ == Tkinter.IntVar:
                    element.set(row[q])

            #Read SQL files for each row
            self.read_sql(id=i)


    def sql_choose(self,id,filename=None):
        """Select SQL file by setting path"""
        #Passed filename - no selector needed
        if filename:
            filepath=os.path.join(os.getcwd(),filename)
        #use fileselector r
        else:
            filepath=tkFileDialog.askopenfilename(filetypes=[("SQL",".sql"),("All",".*")]
            ,title="Select SQL Query "+str(id))
        if filepath:
            self.qr[id]["SQL Query"].delete(0,Tkinter.END)
            self.qr[id]["SQL Query"].insert(0,filepath)
            self.qr[id]["SQL Query"].xview(20)#supposed to shift view to end of text (doesn't work)
            self.read_sql(id)


    def read_sql(self,id):
        """"Read SQL files and stores each in self.qr[id]["sql"] """
        filepath=self.qr[id]["SQL Query"].get()
        if filepath:
            print id,"Opening",filepath
            try:
                sql_file=open(filepath)
            except IOError as e:
                tkMessageBox.showwarning("Load SQL","Failed to open file:"+filepath)
            else:
                self.qr[id]["sql"]=sql_file.read()
                sql_file.close()
        else:
            self.qr[id]["sql"]=""

    def paramupdate(self,*args):
        """Callback for number of parameters field
        (indirectly by associated stringvar)
        Creates number of parameter entry fields in
        accordance with the number from the drop-down
        list.
        """
        #Read the number of fields required
        n=int(self.paramvar.get())

        #Clear any old parameters
        for e in self.param_entries+self.param_entry_names:
            e.destroy()


        self.param_entry_names=[]
        self.param_entries=[]

        #Set new ones
        for i in range(0,n):
            #Entries for parameter headings
            e1=ttk.Entry(self.paramframe, font =("Arial",8, "bold"))
            e1.grid(row=0,column=3+i)
            self.param_entry_names.append(e1)
            #Entries for parameter values
            e2=ttk.Entry(self.paramframe)
            e2.grid(row=1,column=3+i)
            self.param_entries.append(e2)


class TestRunGui:
    """Master part of GUI for running SQL queries"""
    def __init__(self,win_title="Query Thingy"):

        #Set window
        self.win = Tkinter.Tk()
        self.win.title(win_title)
        ##self.win.geometry('800x600')


        #Set grid sizing behaviour (scroll bars don't work without this)
        self.win.grid_rowconfigure(2, weight=1)#Note row 2 is chosen here as we don't want incframe, runframe uncluded
        self.win.grid_columnconfigure(0, weight=1)


        incframe=ttk.Frame(self.win,borderwidth=6, relief="groove")
        incframe.grid(sticky="W")

        #Create Query Definition panel
        self.qp = QueryPanel(frame=incframe)
        self.qp.add_query()#add one row to query panel

        #Run Queries section of screen
        runframe=ttk.Frame(self.win,borderwidth=8, relief="groove")
        runframe.grid(sticky="W",pady=2)#


        title_label = ttk.Label(runframe,text="Run Queries",font =("Arial",8, "bold") )
        title_label.grid(row=0,column=0,sticky="W")

        go_button=ttk.Button(runframe,text="Start", command=self.go)
        go_button.grid(row=1,column=0,pady=0,sticky="W",padx=0)#,columnspan=10)

        #Export Results button
        ex_button=ttk.Button(runframe,text='Export Results',command=self.export_results)
        ex_button.grid(row=1,column=1,sticky="W",padx=12)


        #Radio buttons for export format
        self.rbv=Tkinter.StringVar()#variable associated with radio buttons
        self.rbv.set("txt")#default value
        rb1=Tkinter.Radiobutton(runframe,text="HTML",variable=self.rbv,value="html")
        rb1.grid(row=1,column=2,sticky="w",padx=1)
        rb2=Tkinter.Radiobutton(runframe,text="TSV",variable=self.rbv,value="tsv")
        rb2.grid(row=1,column=3,sticky="w",padx=1)
        rb3=Tkinter.Radiobutton(runframe,text="Text",variable=self.rbv,value="txt")
        rb3.grid(row=1,column=4,sticky="w",padx=1)

        #Create Canvas(needed for scrollbars to work)
        self.canvas=Tkinter.Canvas(self.win)
        self.canvas.grid(sticky='nsew')

        #Create x and y scrollbars. Set their commands to canvas xview and yview
        xscroll = Tkinter.Scrollbar(self.win, orient=Tkinter.HORIZONTAL, command=self.canvas.xview)
        xscroll.grid(row=3, column=0, sticky='we')
        yscroll = Tkinter.Scrollbar(self.win, orient=Tkinter.VERTICAL, command=self.canvas.yview)
        yscroll.grid(row=2, column=1, sticky='ns')

        #Link the canvas x and y scrollcommands to the scrollbars
        self.canvas.configure(xscrollcommand=xscroll.set,yscrollcommand=yscroll.set)

        #Frame to show results within scrollable canvas
        self.resframe=ttk.Frame(self.canvas,borderwidth=0, relief="groove")
        self.resframe.grid(sticky="W")

        #Put the frame in the canvas scrollable zone
        self.canvas.create_window(0, 0, window=self.resframe, anchor='nw')

        #List to hold individual results frames held within resframe
        self.results_frames=[]

        #Update scrollable region (needed for scrollbars to respond to changes in window contents)
        self.resframe.update_idletasks()
        self.canvas.configure(scrollregion=(0, 0, self.resframe.winfo_width(),self.resframe.winfo_height()))

        self.win.mainloop()#Start Event Loop


    def go(self):
        """Button callback for Start Button
        Runs SQL queries
        """
        #Iterate over query row data values (these have been read from elements)
        #Construct QueryData() object for each included one and store in
        #queries list

        #Ensure fields have been read
        self.qp.readfields()

        print"\n*New Run*"

        #Read parameter titles (used by self.export_results)
        self.param_heads = [e.get() for e in self.qp.param_entry_names]
        print"\nParameter Titles:",self.param_heads

        #Read parameter values
        self.params = [e.get() for e in self.qp.param_entries]
        print "\nParameters:",self.params

        #Iterate over query rows
        #Construct QueryData object for each included
        self.queries=[]
        for i,r in enumerate(self.qp.qrd):
            print "\nRow",i

           #See if enabled
            if r["tickvar"]:
                #Construct ODBC connection string
                database=r["Database"]
                username=r["Username"]
                password=r["Password"]
                constring="Driver={Microsoft ODBC for Oracle};Server="+database+";Uid="+username+";Pwd="+password

                #Get associated SQL (note this is from self.qp.qr, not self.qp.qrd)
                sql=self.qp.qr[i]["sql"]

                print "Connection:",constring
                print "SQL"
                print sql

                #Add query data object
                title=r["Description"]+" ("+r["Database"]+")"#Construct title from description and database
                self.queries.append(QueryData(sql,constring,title))

            else:
                #skip row
                print "skip"

        #Run the extracted queries

        #Wipe old results frames
        for f in self.results_frames:
            f.destroy()

        #Create new results frames
        self.results_frames=[]
        for i,q in enumerate(self.queries):
            self.results_frames.append(ttk.Frame(self.resframe,relief="groove",borderwidth=6))#create a frame for each query
            self.results_frames[-1].grid(row=i,column=0,sticky="W")#grid each new frame

        currentcon=None#connection string currently in use (also used to indicate whether the database connection was successfull)
        cnxn=None#holds the database connection
        cursor=None#holds the cursor

        for i,q in enumerate(self.queries):

            #if there's no connection or new query has a different connection
            #string from the current we need to make a new connection
            if q.con<>currentcon:
                #if there's an existing connection, clear it
                if currentcon<>None:
                    cursor.close()
                    cnxn.close()
                    currentcon=None
                #try to connect to the database
                try:
                    cnxn=pyodbc.connect(q.con)
                except pyodbc.Error as e:
                    print "Connection failed:",q.con,e[1]
                else:
                    cursor=cnxn.cursor()
                    currentcon=q.con

            if currentcon: q.execute(cursor,self.params)#Execute the query if we have a successful connection

            #Write the results to the GUI. Still done if connection failed
            #because "empty results" message to be shown.
            q.writeframe(self.results_frames[i])

        #Finished running queries - close cursor and connection if they exist.
        if currentcon:
            cursor.close()
            cnxn.close()

        #Update scrollable region (needed for scrollbars to respond to changes in window contents)
        self.resframe.update_idletasks()
        self.canvas.configure(scrollregion=(0, 0, self.resframe.winfo_width(),self.resframe.winfo_height()))


    def export_results(self):
        """Saves the results to a file when the export results button is pressed"""

        #File type mode handling based on radio buttons
        mode=self.rbv.get()
        if  mode=="html":
            default=".html"
            types=[('HTML Table','.html'),('Any','*')]
            heading="Export results as HTML table(s)"
        elif mode=="tsv":
            default=".tsv"
            types=[('Tab separated','.tsv'),('Any','*')]
            heading="Export results as tab-separated file"
        elif mode=="txt":
            default=".txt"
            types=[('Text','.txt'),('Any','*')]
            heading="Export results as text file"
        else:
            default="*"
            types=[('Any','*')]
            heading="No file format set. This isn't supposed to happen."


        savefile=tkFileDialog.asksaveasfilename(defaultextension=default,filetypes=types,title=heading)
        #If a filename was chosen, loop through all the query results and save the details for each using the query object's makehtml method
        if savefile<>"":

            #css for HTML file mode
            css="""
            <style>
            h1 {font-size:200%;background-color:#111131;margin-bottom:3;margin-top:0;color:#EEEE22}
            h2 {font-size:140%;background-color:#CCCCDC;margin-bottom:8;margin-top:8}
            h3 {font-size:110%;background-color:#111151;margin-bottom:8;margin-top:8;color:#FEFEFE}

            table {border-collapse:collapse;cellspacing:1; cellpadding:1;}
            table, th, td {border:1px solid #555555;}
            th {font-size:100%;font-family:verdana,arial,'sans serif';background-color:#C0C0D0}
            tr {font-size:75%;font-family:verdana,arial,'sans serif';background-color:#FFFFFF;vertical-align:top;}

            body{
            font-size:75%;
            font-family:verdana,arial,'sans serif';
            background-color:#EFEFFF;
            color:#000040;
            margin:10px
            }
            </style>
            """

            #Generate a message used to add the parameter values to the saved file
            if mode=="html":message="<b>Parameters </b><br>"
            else: message="Parameters\n"
            for lab,p in zip(self.param_heads,self.params):
                if mode=="html":message=message+"<i>"+lab+"</i>: "+p+"<br>"
                else: message=message+lab+": "+p+" "
            if mode in("tsv","txt"):
                message=message+"\n"

            #Generate report for each of the queries
            for i,q in enumerate(self.queries):
                if mode=="html":
                    if i==0:q.makehtml(savefile,'n',message,css)#for first query in sequence create a new save file and include the paramerter message
                    else:q.makehtml(savefile,'y')#for subsequent queries in the sequence, append to the first file
                elif mode=="tsv":
                    if i==0:q.makesv(savefile,'n',message)
                    else:q.makesv(savefile,'y')
                elif mode=="txt":
                    if i==0:q.makelongfile(savefile,'n',message)
                    else:q.makelongfile(savefile,'y')


class QueryData:
    """Executes SQL query using PYODBC and stores the results

    Arguments:
        query - string containing the SQL query. Can accept parameters if "?"
        placed in the text.
        con - odbc connection string
        titletext - optional string containing title wording

    """
    def __init__(self,query,con,titletext="No title"):
        self.query=query#string holding the SQL query
        self.con=con#connection string. Not actually used within this class -just stored here for convenience
        self.title=titletext#Optional title text
        self.clear()#ititialise the variables that hold info returned by the query


        #Some ttk Styles for GUI
        style = ttk.Style()
        style.configure("BW.TLabel",foreground="black",background="white", relief=Tkinter.GROOVE)
        style.configure("BW.TEntry",foreground="black", background="white",fieldbackground="white", relief=Tkinter.FLAT)
        style.configure("WB.TEntry",foreground="white", background="black",fieldbackground="black", relief=Tkinter.FLAT)


    def clear(self):
        """Set the variables that hold info returned by the query.
        Also wipes any existing data(needed  for re-queries)"""
        self.headings=[]#holds column headings returned by query
        self.colwidths=[]#holds column widths for formatting purposes
        self.rows=[]#holds PYDOBC row objects for each returned row
        self.runtime=""#holds the execution time and date

    def execute(self,cursor,param=[]):
        """Executes the query (with optional parameters) and stores results in
           self.rows.Requires an existing cursor created outside this class.
           Parameters will only be used if there is a corresponding ? in the SQL

            Arguments:
                cursor - cursor created by PYODBC
                param - (optional) list of  parameters to pass to the query
        """
        self.clear()#wipe old results

        #Execute the query. If there are ?s in the script we should pass the parameter(s)
        #from the parameters list. Slice the param list in case it contains more elements than needed
        pcount=self.query.count("?")#count number of parameters needed by the script
        if pcount==0:cursor.execute(self.query)
        else: cursor.execute(self.query,param[:pcount])

        self.runtime=time.strftime("%d-%b-%Y, %H:%M:%S")#Record the exectuion time/date
        self.rows=cursor.fetchall()#Store all the results as a list of pyodbc row objects
        #extract column names from query
        self.headings=[d[0] for d in cursor.description]
        #set initial values for column widths, based on heading width
        self.colwidths=[len(str(h)) for h in self.headings]

        #loop through rows to adjust column widths to accomodate widest value in results
        for resultrow in(self.rows):#would have called this "row" but it could be confused with the row parameter used by grid
            for c,col in enumerate(resultrow):
                itemwidth=len(str(col))
                if self.colwidths[c]<itemwidth: self.colwidths[c]=itemwidth

    def writeframe(self,frame):
        """ Write the results to an existing tkinter frame"""
        #Wipe any existing contents of the frame
        for w in frame.children.values(): w.destroy()

        #Title Text
        title=ttk.Label(frame, text=self.title+" ("+self.runtime+")", font=("Helvetica",8,"bold"))#Title displayed using text lable
        title.grid(row=0,column=0,sticky="W",columnspan=len(self.headings)+1)#Columnspan needed to stop wide title disrupting the grid (+1 needed as must be >0)

         #Heading Text
        for h,head in enumerate(self.headings):
            fon=("Helvetica",8,"bold");b='black';f='white'#set bold font and white text on black background for heading row
            ##e=ttk.Entry(frame,font=fon, style="WB.TEntry")#create Entry widget to display the column heading
            e=Tkinter.Entry(frame,font=fon,relief=Tkinter.GROOVE,bg=b,fg=f)
            e.grid(row=1,column=h,sticky="W")
            e.config(width=self.colwidths[h]+3)#Adjust the size of the widget to the width of it's contents (the +3 is a fudge factor, seems to he
            e.insert(0, head)#write the heading to the widget

        #Data rows
        fon=("Helvetica",8,"normal");b='white';f='black'#set normal font and black text on white background for remaining text
        for r,resultrow in enumerate(self.rows):#loop through all the data rows over chosen start-end range
            for c,(col,cwidth) in enumerate(zip(resultrow,self.colwidths)):#loop through the row contents and corresponding collumn widths
                e=Tkinter.Entry(frame,font=fon,width=cwidth+3,relief=Tkinter.GROOVE,bg=b,fg=f)
                ##e=ttk.Entry(frame,font=fon,width=cwidth+3,style="BW.TEntry")
                e.grid(row=r+2,column=c,sticky="W")#r+2 because row 0 is used for title text and row 1 for column headings
                e.insert(0, str(col))

        #Print message if no results returned
        if len(self.rows)==0:
            e=ttk.Entry(frame,font=("Helvetica",8,"bold"))#create widget to show message
            e.insert(0,"No Results")
            e.grid(columnspan=1+len(self.headings),sticky="WE")


    def colwriteframe(self,frame):
        """Writes results in frame with headings in a and results in columns
        rather than rows, e.g

        Forename Tim
        Surname  Madeupname
        DOB      01/07/1970

        """
        #Wipe any existing contents of the frame
        for w in frame.children.values(): w.destroy()

        #Headings
        tw=len(max(self.headings,key=len))#width of text box
        heads=Text(frame,font=("Courier",8),height=1,width=tw)
        heads.grid(row=0,column=0,sticky=N)

        for h,head in enumerate(self.headings):
            heads.insert("end",str(head))
            heads.insert("end","\n")
        heads.config(height=h+1)

        #Data
        #determine required width of text box, based on column width values if not null
        if len(self.colwidths)>0:tw=max(self.colwidths)
        else: tw=5

        text=Text(frame,font=("Courier",8),height=2,width=tw)
        text.grid(row=0,column=1,sticky=N)

        for r,resultrow in enumerate(self.rows):#loop through all the data rows over chosen start-end range
            for c,col  in enumerate(resultrow):
                text.insert("end",str(col))
                text.insert("end","\n")
            text.config(height=c+1)#adjust height of the text box

    def show(self):
        """ Prints the results to the console"""
        print self.title
        if self.headings==[]:
            print"No Results"
            return
        #Print headings
        for i,h in  enumerate(self.headings):
            print h.ljust(self.colwidths[i]),
        print""
        #Print row data
        for r in self.rows:
            for i,c in enumerate(r):print str(c).ljust(self.colwidths[i]),
            print""
        print""

    def makehtml(self,filename="",append='n',message="",css="",start=0,end=None):
        """Writes the results to a file as an HTML table """
        #set file write mode to append or write
        if append=='y':mode='a'
        else: mode='w'
        with open(filename,mode) as outfile:
            #Write optional message text
            if css:outfile.write(css)
            if message:outfile.write("\n<p>"+message+"</p>\n")
            #Write title line
            db=self.con.split(";")[1][7:]#extract database name from the connection string
            outfile.write("<h3>"+self.title+" - "+db+" ("+self.runtime+")"+"</h3>\n")
            #If there aren't any results write a message saying so
            if  self.rows==[]:
                outfile.write("\n<p><i>No results</i></p>\n")
            else:
                #write html tag for start of table
                outfile.write('<table border="1" cellspacing=1 cellpadding=1>\n\n')
                #Headings row
                outfile.write('<tr>\n')
                for h in self.headings:
                    outfile.write('<th>'+h+'</th>'+"\n")
                outfile.write("</tr>\n\n")#end of headings row
                #Data rows
                for r in self.rows[start:end]:#loop through all the data rows over chosen start-end range
                    outfile.write('<tr>\n')#Start new row in the html table]
                    #write each field in the curent data row to the html row
                    for c in r:
                        outfile.write('<td>'+str(c)+'</td>'+"\n")
                    outfile.write("</tr>\n\n")#end of html table row
                outfile.write("</table>\n<br>\n\n")#end of html table

    def makesv(self,filename="",append='n',message="",start=0,end=None):
        """Writes the results to a tab-separated variable file """
        #set file write mode to append or write
        if append=='y':mode='a'
        else: mode='w'
        with open(filename,mode) as outfile:
            #Write optional message text
            if message:outfile.write(message+"\n")
            #Write title line
            db=self.con.split(";")[1][7:]#extract database name from the connection string
            outfile.write(self.title+" - "+db+" ("+self.runtime+")"+"\n")
            #If there aren't any results write a message saying so
            if  self.rows==[]:
                outfile.write("No results\n")
            else:
                #Headings row
                for h in self.headings:
                    outfile.write(h+"\t")
                outfile.write("\n")#end of headings row
                #Data rows
                for r in self.rows[start:end]:#loop through all the data rows over chosen start-end range
                    #write each field in the curent data row to the file
                    for c in r:
                        outfile.write(str(c)+"\t")
                    outfile.write("\n")#end of row
                outfile.write("\n")#end

    def makelongfile(self,filename="",append='n',message="",start=0,end=None):
        """Writes the results to single column text file """
        #set file write mode to append or write
        if append=='y':mode='a'
        else: mode='w'
        with open(filename,mode) as outfile:
            #Write optional message text
            if message:outfile.write(message+"\n")
            #Write title line
            db=self.con.split(";")[1][7:]#extract database name from the connection string
            outfile.write(self.title+" - "+db+" ("+self.runtime+")"+"\n")
            #If there aren't any results write a message saying so
            if  self.rows==[]:
                outfile.write("No results\n\n")
            else:
                #Headings row
                ##outfile.write("New Thing"+"\n")
                #Data rows
                for i,r in enumerate(self.rows[start:end]):#loop through all the data rows over chosen start-end range
                    #write each field in the curent data row to
                    outfile.write("\nRow "+str(i)+"\n")
                    for j,c in enumerate(r):
                        outfile.write(self.headings[j]+": ")
                        outfile.write(str(c)+"\n")
                    outfile.write("\n")#end of row
                outfile.write("\n")#end

    def getcolumn(self,col=0):
        """Retrieves one column from the results and returns it as a list """
        #return if there is no data or the value of col is too high or low
        if self.rows==[]:return
        if col<0: return
        if col+1>len(self.rows):return
        #Extract the data for the chosen row and return it
        templist=[]
        for r in self.rows:
            templist.append(r[col])
        return templist




if __name__=="__main__":
    go = TestRunGui("SQL Query Set")
