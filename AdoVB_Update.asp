<%@ LANGUAGE = VBScript %>
<%  Option Explicit		
    Response.Buffer = true
%>
<HTML>
<!--

* ----------------------------------------------------------------------      
*                                                                             
*                  Synergy - Synergy Language Version 7                       
*                                                                             
*                            Copyright (C) 2001
*     by Synergex International Corporation.  All rights reserved.            
*                                                                             
*         May not be copied or disclosed without the permission of            
*                 Synergex International Corporation                          
*                                                                             
* -----------------------------------------------------------------------     
* -----------------------------------------------------------------------     
*                                                                             
* Source:     AdoVB_Update.asp
*                                                                             
* Facility:   Example for using Microsoft Universal Data Access components 
*	        from an ASP page with VBScript to update the Plants 
*               example database installed with SynergyDE SQL Connectivity.             
*                                                                             
* Abstract:   Displays the before and after values to show the update
*               occurred.
*                                                                             
*             You may change the connect string and/or SQL command            
*               as needed.                                                      
*                                                                             
* $Revision:     $                                                            
*                                                                             
* $Date:         $                                                            
*                                                                             
--------------------------------------------------------------------------    
-->
<HEAD>
    <TITLE>Simple ADO Update with ASP Using VBScript</TITLE>
<%

' Globals

Const adStateClosed = &H00000000 ' From ADOVBS.INC

Dim status
    status = 0
    if (status = 0) then call execute_rs

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' execute_rs()
' Using ADO, connect to the database and create an HTML table with
' the recordset
'
sub execute_rs

    Dim oConn
    Dim oRs
    Dim ix

    ' Display information for the user      
    
    Response.Write("<font size=""4"" face=""Arial, Helvetica"">")
    Response.Write("<b>Simple ADO Query with ASP Using VBScript</b></font><br>")
    Response.Write("<hr size=""1"" color=""#000000"">List of Available Plants:<br><br></hr>")
    
    ' Forward any errors to the CheckADOError() error handler
    
    on error resume next
      
    ' Instantiate the ADO objects. Objects must always be instantiated.
      
    Set oConn = Server.CreateObject("ADODB.Connection")
    Set oRs = Server.CreateObject("ADODB.Recordset")
      
    ' Open the connection
          
    oConn.Open "DSN=xfODBC;UID=DBADMIN;PWD=MANAGER;DBQ=sodbc_sa;"
    CheckADOError oConn,status

    if (status <> 0) then
        ' Failure encountered, make sure to close any recordset and/or
        ' connection then set the object variable to nothing
        if (oConn.state <> adStateClosed) then oConn.close
        set oRs = nothing
        set oConn = nothing
        exit sub
    end if 

    ' ~~~ BEFORE UPDATE ~~~
              
    ' Execute the SQL query and output the recordset
      
    Set oRs = oConn.Execute("SELECT in_itemid, in_name, in_price FROM public.plants WHERE {fn LCASE(in_name)} LIKE 'sour gum%'")
    CheckADOError oConn,status
    
    if (status <> 0) then 
        if (oRs.State <> adStateClosed) then oRs.Close
        if (oConn.State <> adStateClosed) then oConn.Close
        set oRs = nothing
        set oConn = nothing
        exit sub
    end if 

    ' Send the client a recordset as a table
          
    Response.Write("<br><hr></hr><h2>Before Update</h2><br><TABLE border = 1>")
      
    Do while (Not oRs.eof)
        Response.Write("<tr>")
        For ix=0 to (oRs.fields.count-1)
            Response.Write("<TD VAlign=top>")
            Response.Write(oRs(ix))
            Response.Write("</TD>")
        Next
        Response.Write("</tr>")
         
        oRs.MoveNext
        CheckADOError oConn,status
        
        if (status <> 0) then 
            if (oRs.State <> adStateClosed) then oRs.Close
            if (oConn.State <> adStateClosed) then oConn.Close
            set oRs = nothing
            set oConn = nothing
	    exit sub
        end if 
    Loop
      
    Response.Write("</TABLE>")

    ' ~~~ UPDATE PROCESS ~~~
    
    ' Perform the Update operation
    
    Response.Write("<br><hr></hr><h2>Update - Sour Gum price increase of 10%</h2><br><TABLE border = 1>")
    oConn.Execute ("UPDATE public.plants SET in_price = in_price + (in_price * .10) WHERE {fn LCASE(in_name)} LIKE 'sour gum%'")
    CheckADOError oConn,status
    
    if (status <> 0) then 
        if (oRs.State <> adStateClosed) then oRs.Close
        if (oConn.State <> adStateClosed) then oConn.Close
        set oRs = nothing
        set oConn = nothing
        exit sub
    end if 
    
    ' ~~~ AFTER UPDATE ~~~
    
    ' Execute the SQL query and output the updated values
              
    Set oRs = oConn.Execute("SELECT in_itemid, in_name, in_price FROM public.plants WHERE {fn LCASE(in_name)} LIKE 'sour gum%'")
    CheckADOError oConn,status
    
    if (status <> 0) then 
        if (oRs.State <> adStateClosed) then oRs.Close
        if (oConn.State <> adStateClosed) then oConn.Close
        set oRs = nothing
        set oConn = nothing
        exit sub
    end if 

    ' Send the client a recordset as a table

    Response.Write("<br><hr></hr><h2>After Update</h2><br><TABLE border = 1>")
      
    Do while (Not oRs.eof)
        Response.Write("<tr>")
        For ix=0 to (oRs.fields.count-1)
            Response.Write("<TD VAlign=top>")
            Response.Write(oRs(ix))
            Response.Write("</TD>")
        Next
        Response.Write("</tr>")
         
        oRs.MoveNext
        CheckADOError oConn,status
        if (status <> 0) then 
            if (oRs.State <> adStateClosed) then oRs.Close
            if (oConn.State <> adStateClosed) then oConn.Close
            set oRs = nothing
            set oConn = nothing
	    exit sub
        end if 
    Loop
      
    Response.Write("</TABLE>")
      
    ' Release resources by closing the result-set and connection then set the
    ' objects to nothing.
          
    if (oRs.State <> adStateClosed) then oRs.Close
    if (oConn.State <> adStateClosed) then oConn.Close
    set oRs = nothing
    set oConn = nothing
    
end sub

' ~~ CheckADOError() ~~~~~~~~~~~~~~~~~
' Call this routine after every ADO operation
' to neatly catch errors and clear the error messages
'
Sub CheckADOError(objConnection,connStatus)
    dim oErr      

    ' Continue if no errors encoutered.
    
    if (Err.number = 0) then exit sub

    ' Clear the response buffer and send only the error message
    ' to the client
        
    response.clear
    response.write("<html><head></head><body><h1>An application error occurred</h1><br><hr size=5>")
    if (objConnection.Errors.Count <> 0) then
        for each oErr In objConnection.Errors
           Response.Write("ADO Error #" & oErr.Number & "<BR>")
           Response.Write("  " & oErr.Description & "<BR>")
           Response.Write("  SQL State  :" & oErr.SQLState & "<BR>")
           Response.Write("  NativeError:" & oErr.NativeError & "<BR><HR>")
        next
        objConnection.Errors.clear
     end if    
' duplicate container as oErr ... Response.Write("Error # " & CStr(Err.number) & " " & Err.Description)
    Err.Clear
    response.write("</body></html>")
    connStatus = 1
End Sub

%>

</HEAD>

<BODY BGCOLOR="White" topmargin="10" leftmargin="10">
</BODY>
</HTML>
