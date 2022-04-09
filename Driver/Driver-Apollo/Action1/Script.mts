Dim karr
Sarray=Array("s","s","s","s","s","s","s","s","s","s") 'Passed the Keyword trough an array
Dim index,flag
index=0
Set objexcel=CreateObject("Excel.Application")
'Reads Data from the Excel Sheet whose path is mentioned below.
Set objworkbook=objexcel.Workbooks.open("C:\Users\sfjbs\Desktop\HybridFramework\Organizer\OrganizerSheet_Shalini.xlsx")
Set objmodulesheet=objworkbook.Worksheets(1)
modrowcount=objmodulesheet.UsedRange.Rows.Count
msgbox modrowcount

Set objtestcasesheet=objworkbook.Worksheets(2)
tcrowcount=objtestcasesheet.UsedRange.Rows.Count
msgbox tcrowcount

Set objteststepsheet=objworkbook.Worksheets(3)
tsrowcount=objteststepsheet.UsedRange.Rows.Count
msgbox tsrowcount

Services.StartTransaction "AppointmentBooking_Tr1"

For i=1 to modrowcount Step 1
	

	modexe=objmodulesheet.cells(i,3)
	If modexe="Y" Then
	
	moduleid=objmodulesheet.cells(i,1)
	
		For j = 1 To tcrowcount Step 1
		
			tmoduleid=objtestcasesheet.cells(j,5)
			
			If moduleid=tmoduleid and objtestcasesheet.cells(j,4)="Yes" Then
				
				tc_testcaseid=objtestcasesheet.cells(j,1)
				
				For k = 1 To tsrowcount Step 1
					ts_testcaseid=objteststepsheet.cells(k,5)
						If tc_testcaseid=ts_testcaseid Then
							keyword=objteststepsheet.cells(k,4)
							
							flag=0
							
							For p=0 To 9 Step 1
									If Sarray(p)=keyword Then
										flag=1
										objteststepsheet.cells(k,7)="Executed"
									End If
									Next
									
									If flag=0 Then
										Sarray(index)=keyword
										index=index+1
										
										
									Select Case keyword
'This case will allow the user to Add any number of patients they want to by providing details such as: First Name, Last Name, Date Of Birth, Relation.		
                                                                        Case "ANP"
										msgbox "Add New Patient Function will run"
										
									Var1=objteststepsheet.cells(k,6)
									Split1=Split(Var1,":")
								        ANP Split1(1),Split1(3),Split1(5)
									objteststepsheet.cells(k,7)="Executed"

'This case will just check the navigation of Appointmnet Booking Page that is wheather it is clickable or not and after clicking does it navigates to the default page.								
										Case "ABO"
										msgbox "Function Appointment Booking Option will run"
										Call  ABO()
										objteststepsheet.cells(k,7)="Executed"
'This case will check wheather Search Option is enable or not, can a user search doctor, Specialist by their name apart from the default list present										
										Case "SC"
										msgbox "Function Search Check will run"
										Call SC()
									objteststepsheet.cells(k,7)="Executed"
'This case will check the page navigation function.										
										Case "PC"
										msgbox "Function Page Check will run"
									        Call PC()
									      objteststepsheet.cells(k,7)="Executed"  
'This case will check all the filter present and also we will be applying one Filter called (Experience ) in order to check wheather the list of doctors displayed are according to the filter applied.									       
									       Case "FO"
										msgbox "Function Filter Option will run"
										Call FO()
									objteststepsheet.cells(k,7)="Executed"	
'This Case will check all the details that are defined by the tester during Manual testing are present or not along with checking the Hospital Visit Option.
										Case "DD"
										msgbox "Function Doctor Detail will run"
										Call DD()
									objteststepsheet.cells(k,7)="Executed"	
'This case will allow the user to book the prefered slot :1) Slot for Today's Date 2) Slot for Tomorrow's date 3) Slot for any random date within 7 days. Here I am booking the slot for Dr. Manish Pendse.									
									       Case "SLC"
										msgbox "Slot Check Function will run"
										Call SLC()
									objteststepsheet.cells(k,7)="Executed"	
									
'This will validate wheather Consultation Charges can be seen on the Appointment page before confirming the slot.	                                                                   
										Case "SP"
										msgbox "Select Patient Function will run"
										Call SP()
									objteststepsheet.cells(k,7)="Executed"
									
'This case will check wheather a user can select the patient that wants to take the consultation 
                                                                         Case "CCC"
                                                                         msgbox "Checking consultation charges"
                                                                         Call CCC()
                                                                      objteststepsheet.cells(k,7)="Executed"
				
' This case will validate wheather user get any kind of Pop-up message post confirmation. Also it check wheather the user can Reschudule or Cancel his/her appointmnet.										
										Case "SCO"
										msgbox "Function Slot Confirmation will run"
										Call SCO()
									objteststepsheet.cells(k,7)="Executed"
									
									End Select	
									End If
							
						End If
				Next
			End If
			
		Next
		
	End If
	
Next


Services.EndTransaction "AppointmentBooking_Tr1"

objexcel.quit
Set objexcel=nothing
Set objworkbook=nothing
Set objmodulesheet=nothing
Set objtestcasesheet=nothing
Set objteststepsheet=nothing @@ script infofile_;_ZIP::ssf110.xml_;_

 @@ script infofile_;_ZIP::ssf112.xml_;_

 @@ script infofile_;_ZIP::ssf115.xml_;_
