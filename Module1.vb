Imports System.Text
Imports System.Convert
Imports System.Math
Imports System
Imports System.IO
Imports System.Threading

Imports VISSIM_COMSERVERLib

Module Module1
    Dim vissim As Vissim
    Dim net As Net
    Dim simulation As Simulation
    Dim vehicles As Vehicles
    Dim vehicle As Vehicle
    Dim links As Links
    Dim link As Link
    Dim evaluation As Evaluation
    Dim simtime = 10800
    Dim resolution = 5

    'Initiating function flipcoin to generate binary output that depends on the probability as its input parameter
    Function flipcoin(ByVal j As Integer) As Integer
        If j = 0 Then
            Return 1

        ElseIf j = 100 Then
            Return 0

        Else
            Dim i As Integer = CInt(Math.Floor((100 - 0 + 1) * Rnd())) + 0
            If i <= j Then
                Return 0
            Else
                Return 1
            End If
        End If

    End Function

    Sub Main()
        'For loop for different random seeds
        For randcount = 1 To 2
            Dim randseed = randcount

            'Initiating an array named as "vol" with volume inputs accessed by "Volume" on eastbound and westbound approaches for the second hour 
            'Running a for loop to access all the volumes
            Dim vol() As Double = {1000, 2400}
            For Volume As Integer = 0 To vol.Length - 1

                'Initiating an array named as "vol" with volume inputs accessed by "Volume" on eastbound and westbound approaches for the second hour 
                'Running a for loop to access all the likelihoods
                Dim likelihood() As Integer = {0, 20, 50, 90}
                For value As Integer = 0 To likelihood.Length - 1

                    'Initializing a new instance of Vissim
                    vissim = New Vissim

                    'loading Vissim network model ".inp" file and ".ini" file - DOS platform
                    vissim.LoadNet("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.inp")
                    vissim.LoadLayout("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.ini")
                    'minimizing the vissim model to make it run faster
                    vissim.ShowMinimized()

                    'Initiating vehicle input variable to provide vehicle input in the Vissim model 
                    'vehinps - colletion of various vehicle inputs
                    Dim vehinps = vissim.Net.VehicleInputs

                    'Initiating vehicle input for all approached: Eastbound, Westbound, Northbound, Southbound
                    Dim vehinpEB = vissim.Net.VehicleInputs.GetVehicleInputByNumber(1)
                    Dim vehinpNB = vissim.Net.VehicleInputs.GetVehicleInputByNumber(2)
                    Dim vehinpSB = vissim.Net.VehicleInputs.GetVehicleInputByNumber(3)
                    Dim vehinpWB = vissim.Net.VehicleInputs.GetVehicleInputByNumber(4)

                    'AddVehiclInput to add a new vehicle input for Eastbound approach for first hour of the run and 
                    'setting up it's attribute value volume and traffic composition
                    vehinpEB = vehinps.AddVehicleInput(1, 0, 3600)
                    vehinpEB.AttValue("Volume") = 1800
                    vehinpEB.AttValue("TrafficComposition") = 4

                    'Initiating variables to watch first hour volume, time until and time from for Eastbound-first hour
                    Dim outputvolEB1 As Integer
                    Dim outputtimeuntilEB1 As Integer
                    Dim outputtimefromEB1 As Integer

                    'Assigning the variables with the first hour volume, time until and time from 
                    outputvolEB1 = vehinpEB.AttValue("Volume")
                    outputtimeuntilEB1 = vehinpEB.AttValue("TimeUntil")
                    outputtimefromEB1 = vehinpEB.AttValue("TimeFrom")

                    'AddVehiclInput to add a new vehicle input for Westbound approach for first hour of the run and 
                    'setting up it's attribute value volume and traffic composition
                    vehinpWB = vehinps.AddVehicleInput(5, 0, 3600)
                    vehinpWB.AttValue("Volume") = 1800
                    vehinpWB.AttValue("TrafficComposition") = 3

                    'Initiating variables to watch first hour volume, time until and time from for Westbound-first hour
                    Dim outputvolWB1 As Integer
                    Dim outputtimeuntilWB1 As Integer
                    Dim outputtimefromWB1 As Integer

                    'AddVehiclInput to add a new vehicle input for Westbound approach for first hour of the run and 
                    'setting up it's attribute value volume and traffic composition
                    outputvolWB1 = vehinpWB.AttValue("Volume")
                    outputtimeuntilWB1 = vehinpWB.AttValue("TimeUntil")
                    outputtimefromWB1 = vehinpWB.AttValue("TimeFrom")

                    'activating evaluation tools in the file
                    evaluation = vissim.Evaluation
                    evaluation.AttValue("traveltime") = True
                    evaluation.AttValue("delay") = True
                    evaluation.AttValue("datacollection") = True

                    'Assigining simulation properties
                    simulation = vissim.Simulation
                    simulation.Period = simtime
                    simulation.Resolution = resolution
                    simulation.RandomSeed = randseed


                    'Initializing the intersection link as blocking box and the link zone where vehicles are assigned
                    'with blocking or non-blocking type as decisionBox for Eastbound and Westbound approaches
                    Dim decisionBoxEB = vissim.Net.Links.GetLinkByNumber(11)
                    Dim blockingBoxEB = vissim.Net.Links.GetLinkByNumber(13)
                    Dim decisionBoxWB = vissim.Net.Links.GetLinkByNumber(6)
                    Dim blockingBoxWB = vissim.Net.Links.GetLinkByNumber(14)

                    'Initializing the variables to get vehicles in the decison box and blocking box
                    Dim decisionBoxVehiclesEB = decisionBoxEB.GetVehicles
                    Dim blockingBoxVehiclesEB = blockingBoxEB.GetVehicles
                    Dim decisionBoxVehiclesWB = decisionBoxWB.GetVehicles
                    Dim blockingBoxVehiclesWB = blockingBoxWB.GetVehicles

                    'Initializing the two flags to use them for logic test in the if commands for changing the volumes  
                    'during second and thrid hour of Eastbound and Westbound 
                    Dim flag As Boolean = False
                    Dim flag2 As Boolean = False

                    'For loop for timesteps 
                    'Run the simulation as singlestep
                    For j = 0 To (simtime - 1) * resolution
                        simulation.RunSingleStep()

                        'Add new vehicle input with new attributes of volume, timeuntil and time from for second hour of simulation
                        'for eastbound and westbound

                        If (flag = False) And (simulation.AttValue("elapsedtime") > 3600) Then

                            'Adding vehicle input and setting it's attribute values for eastbound for second hour
                            vehinpEB = vehinps.AddVehicleInput(1, 3601, 7200)
                            vehinpEB.AttValue("Volume") = 3000
                            vehinpEB.AttValue("TrafficComposition") = 4

                            'Initiating variables to watch second hour vehicle input attributes for Eastbound
                            Dim outputvolEB2 As Integer
                            Dim outputtimeuntilEB2 As Integer
                            Dim outputtimefromEB2 As Integer

                            'Assigning the variables created to watch second hour vehicle input attributes with values for Eastbound
                            outputvolEB2 = vehinpEB.AttValue("Volume")
                            outputtimeuntilEB2 = vehinpEB.AttValue("TimeUntil")
                            outputtimefromEB2 = vehinpEB.AttValue("TimeFrom")

                            'Adding vehicle input and setting it's attribute values for westbound for second hour
                            vehinpWB = vehinps.AddVehicleInput(5, 3601, 7200)
                            vehinpWB.AttValue("Volume") = 2600
                            vehinpWB.AttValue("TrafficComposition") = 3

                            'Initiating variables to watch second hour vehicle input attributes for Westbound
                            Dim outputvolWB2 As Integer
                            Dim outputtimeuntilWB2 As Integer
                            Dim outputtimefromWB2 As Integer

                            'Assigning the variables created to watch second hour vehicle input attributes with values for westbound
                            outputvolWB2 = vehinpWB.AttValue("Volume")
                            outputtimeuntilWB2 = vehinpWB.AttValue("TimeUntil")
                            outputtimefromWB2 = vehinpWB.AttValue("TimeFrom")

                            flag = True
                        End If


                        'Add new vehicle input with new attributes of volume, timeuntil and time from for third hour of simulation
                        'for eastbound and westbound

                        If flag2 = False And simulation.AttValue("elapsedtime") > 7200 Then

                            'Adding vehicle input and setting it's attribute values for eastbound for third hour
                            vehinpEB = vehinps.AddVehicleInput(1, 7201, 10800)
                            vehinpEB.AttValue("Volume") = 1700
                            vehinpEB.AttValue("TrafficComposition") = 4

                            'Initiating variables to watch third hour vehicle input attributes for Eastbound
                            Dim outputvol3 As Integer
                            Dim outputtimeuntil3 As Integer
                            Dim outputtimefrom3 As Integer

                            'Assigning the variables created to watch third hour vehicle input attributes with values for Eastbound
                            outputvol3 = vehinpEB.AttValue("Volume")
                            outputtimeuntil3 = vehinpEB.AttValue("TimeUntil")
                            outputtimefrom3 = vehinpEB.AttValue("TimeFrom")

                            'Adding vehicle input and setting it's attribute values for westbound for second hour
                            vehinpWB = vehinps.AddVehicleInput(5, 7201, 10800)
                            vehinpWB.AttValue("Volume") = 1500
                            vehinpWB.AttValue("TrafficComposition") = 3

                            'Initiating variables to watch second hour vehicle input attributes for Westbound
                            Dim outputvolWB3 As Integer
                            Dim outputtimeuntilWB3 As Integer
                            Dim outputtimefromWB3 As Integer

                            'Assigning the variables created to watch second hour vehicle input attributes with values for westbound
                            outputvolWB3 = vehinpWB.AttValue("Volume")
                            outputtimeuntilWB3 = vehinpWB.AttValue("TimeUntil")
                            outputtimefromWB3 = vehinpWB.AttValue("TimeFrom")

                            flag2 = True
                        End If

                        'WB lanewise sampling of leading vehicles as blocking or non-blocking which the follower vehicle also get assigned to

                        'The WB decisionBox and WB blockingBox are checked for presence of vehicle 
                        If blockingBoxVehiclesWB.Count > 0 And decisionBoxVehiclesWB.Count > 0 Then

                            'If there is vehicle present in decisionBox and blockingBox then for each of the three lane "i" as 1, 2,3
                            For i As Integer = 1 To 3
                                'Console.Write(i)

                                ' Initiating variables to store coordinate and index of the vehicle in the blockingbox that is closest
                                ' to the intersection
                                Dim foundblockingBoxMax = 0
                                Dim maxblockingBoxCoord As Double
                                Dim maxblockingBoxIndex As Integer

                                'All vehicles in the blocking box are compared for their coordinates and the one closest to the 
                                'decisionBox is identified and its coordinates and ID is stored. This identification of vehicles
                                'in the blcoking box closest to the decisionBox is happening for every lane individually that is
                                'Lane 1, Lane 2 and Lane 3
                                For q = 1 To blockingBoxVehiclesWB.Count
                                    If (blockingBoxVehiclesWB(q).AttValue("Lane") = i) Then
                                        'Console.WriteLine(blockingBoxVehiclesWB(q).AttValue("POINT").X)
                                        If (foundblockingBoxMax = 0) Then
                                            maxblockingBoxCoord = blockingBoxVehiclesWB(q).AttValue("POINT").X
                                            maxblockingBoxIndex = q
                                            foundblockingBoxMax = 1
                                        ElseIf (blockingBoxVehiclesWB(q).AttValue("POINT").X > maxblockingBoxCoord) Then
                                            maxblockingBoxCoord = blockingBoxVehiclesWB(q).AttValue("POINT").X
                                            maxblockingBoxIndex = q
                                        End If
                                        ' Console.Write(" - ")
                                        ' Console.Write(maxblockingBoxIndex)
                                    End If
                                Next

                                'The vehicle identified in the blockingbox closest to the decisionbox is stored in maxblockingBoxIndex
                                'If the speed of this identified vehicle is less thab 15 km/h then it enters the following loop and the
                                'the vehicles in the decisionBox are checked for lane change, lead vehicle is identified and is 
                                'assigned as blocking or non-blocking vehicle 

                                '## COLOR CODES AND VEHICLE TYPES ##
                                'Blocking Vehicle Type     -> 1
                                'Non-Blcoking Vehicle Type -> 2
                                'Blocking Leader        ->  Purple  ->  RGB(255, 0, 255)
                                'Non-Blocking leader    ->  Orange  ->  RGB(255, 165, 0)
                                'Blocking follower      ->  Black   ->  RGB(0, 0, 0)       -> RGB(255 * vehtype1, 255 * vehtype1, 255 * vehtype1)
                                'Non-Blocking follower  ->  White   ->  RGB(255, 255, 255) -> RGB(255 * vehtype2, 255 * vehtype2, 255 * vehtype2)

                                If (foundblockingBoxMax = 1) Then
                                    ' Console.Write(blockingBoxVehiclesWB(maxblockingBoxIndex).AttValue("Speed"))
                                    If (blockingBoxVehiclesWB(maxblockingBoxIndex).AttValue("Speed") < 15) Then
                                        Dim foundMax1 = 0
                                        Dim minCoord1 As Double
                                        Dim minIndex1 As Integer

                                        'all vehicles in the decisionbox that change lane while the value to block is 0 
                                        'then they get assigned to orange color and remain non-blocking type
                                        For b = 1 To decisionBoxVehiclesWB.Count
                                            If (decisionBoxVehiclesWB(b).AttValue("LANECHANGE") <> 0 And likelihood(value) = 0) Then
                                                decisionBoxVehiclesWB(b).AttValue("Type") = 2
                                                If (decisionBoxVehiclesWB(b).AttValue("Type") = 2) Then
                                                    decisionBoxVehiclesWB(b).AttValue("color") = RGB(255, 165, 0)
                                                End If
                                            End If

                                            'all vehicles in the decisionbox that change lane while the value to block greater than 0 
                                            'then they get assigned to purple color and become blocking type as they show aggresive tendency
                                            If (decisionBoxVehiclesWB(b).AttValue("LANECHANGE") <> 0 And likelihood(value) <> 0) Then
                                                decisionBoxVehiclesWB(b).AttValue("Type") = 1
                                                If (decisionBoxVehiclesWB(b).AttValue("Type") = 1) Then
                                                    decisionBoxVehiclesWB(b).AttValue("color") = RGB(255, 0, 255)
                                                End If
                                            End If
                                        Next

                                        'All the vehicles in the decisioBox are compared to find ID and coordinate of vehicle leading
                                        'in the decision box for each lane
                                        For p = 1 To decisionBoxVehiclesWB.Count
                                            If (decisionBoxVehiclesWB(p).AttValue("Lane") = i) Then
                                                'Console.Write(p)
                                                'Console.Write(" - ")
                                                'Console.WriteLine(decisionBoxVehiclesWB(p).AttValue("POINT").X)
                                                If (foundMax1 = 0) Then
                                                    minCoord1 = decisionBoxVehiclesWB(p).AttValue("POINT").X
                                                    minIndex1 = p
                                                    foundMax1 = 1
                                                ElseIf (decisionBoxVehiclesWB(p).AttValue("POINT").X < minCoord1) Then
                                                    minCoord1 = decisionBoxVehiclesWB(p).AttValue("POINT").X
                                                    minIndex1 = p
                                                End If
                                            End If
                                        Next

                                        'if the vehicle identified as leading in the decision box is unassigned to blocking or non-blocking type
                                        'then it is assigned to blocking purple or non-blocking orange depending on the likelihood value 
                                        'assigned for the run if
                                        If foundMax1 = 1 Then
                                            If (decisionBoxVehiclesWB(minIndex1).AttValue("Type") = 30) Then
                                                decisionBoxVehiclesWB(minIndex1).AttValue("Type") = flipcoin(likelihood(value)) + 1

                                                If (decisionBoxVehiclesWB(minIndex1).AttValue("Type") = 1) Then
                                                    decisionBoxVehiclesWB(minIndex1).AttValue("color") = RGB(255, 0, 255)
                                                End If

                                                If (decisionBoxVehiclesWB(minIndex1).AttValue("Type") = 2) Then
                                                    decisionBoxVehiclesWB(minIndex1).AttValue("color") = RGB(255, 165, 0)
                                                End If
                                            End If

                                            'all the follower vehicles in the decisionbox are assigned to the same type as the lead vehicle 
                                            'and appear as white-non blocking or black-blocking
                                            For k = 1 To decisionBoxVehiclesWB.Count
                                                If (decisionBoxVehiclesWB(k).AttValue("Lane") = i) Then
                                                    Dim vehicle = decisionBoxVehiclesWB(k)
                                                    If Not vehicle Is Nothing Then
                                                        If (k <> minIndex1) Then
                                                            If (decisionBoxVehiclesWB(k).AttValue("Type") = 30) And (decisionBoxVehiclesWB(minIndex1).AttValue("Type") = 1) Then
                                                                decisionBoxVehiclesWB(k).AttValue("Type") = 1
                                                                Dim vehtype1 As Integer = decisionBoxVehiclesWB(k).AttValue("Type") - 1
                                                                decisionBoxVehiclesWB(k).AttValue("color") = RGB(255 * vehtype1, 255 * vehtype1, 255 * vehtype1)

                                                            ElseIf (decisionBoxVehiclesWB(k).AttValue("Type") = 30) And (decisionBoxVehiclesWB(minIndex1).AttValue("Type") = 2) Then
                                                                decisionBoxVehiclesWB(k).AttValue("Type") = 2
                                                                Dim vehtype2 As Integer = decisionBoxVehiclesWB(k).AttValue("Type") - 1
                                                                decisionBoxVehiclesWB(k).AttValue("color") = RGB(255 * vehtype2, 255 * vehtype2, 255 * vehtype2)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next k
                                        End If
                                    End If
                                End If
                            Next
                        End If


                        'EB lanewise sampling of leading vehicles as blocking or non-blocking which the follower vehicle become

                        'The WB decisionBox and WB blockingBox are checked for presence of vehicle 
                        If blockingBoxVehiclesEB.Count > 0 And decisionBoxVehiclesEB.Count > 0 Then

                            'If there is vehicle present in decisionBox and blockingBox then for each of the three lane "n" as 1, 2,3
                            For n As Integer = 1 To 3

                                ' Initiating variables to store coordinate and index of the vehicle in the blockingbox that is closest
                                ' to the intersection
                                Dim foundblockingBoxMin = 0
                                Dim minblockingBoxCoord As Double
                                Dim minblockingBoxIndex As Integer

                                'All vehicles in the blocking box are compared for their coordinates and the one closest to the 
                                'decisionBox is identified and its coordinates and ID is stored. This identification of vehicles
                                'in the blcoking box closest to the decisionBox is happening for every lane individually that is
                                'Lane 1, Lane 2 and Lane 3
                                For r = 1 To blockingBoxVehiclesEB.Count
                                    If (blockingBoxVehiclesEB(r).AttValue("Lane") = n) Then
                                        'Console.Write(r)
                                        'Console.Write(" - ")
                                        'Console.WriteLine(blockingBoxVehiclesEB(r).AttValue("POINT").X)
                                        If (foundblockingBoxMin = 0) Then
                                            minblockingBoxCoord = blockingBoxVehiclesEB(r).AttValue("POINT").X
                                            minblockingBoxIndex = r
                                            foundblockingBoxMin = 1
                                        ElseIf (blockingBoxVehiclesEB(r).AttValue("POINT").X < minblockingBoxCoord) Then
                                            minblockingBoxCoord = blockingBoxVehiclesEB(r).AttValue("POINT").X
                                            minblockingBoxIndex = r
                                        End If
                                    End If
                                Next

                                'The vehicle identified in the blockingbox closest to the decisionbox is stored in maxblockingBoxIndex
                                'If the speed of this identified vehicle is less thab 15 km/h then it enters the following loop and the
                                'the vehicles in the decisionBox are checked for lane change, lead vehicle is identified and is 
                                'assigned as blocking or non-blocking vehicle 

                                '## COLOR CODES AND VEHICLE TYPES ##
                                'Blocking Vehicle Type     -> 1
                                'Non-Blcoking Vehicle Type -> 2
                                'Blocking Leader        ->  Purple  ->  RGB(255, 0, 255)
                                'Non-Blocking leader    ->  Orange  ->  RGB(255, 165, 0)
                                'Blocking follower      ->  Black   ->  RGB(0, 0, 0)       -> RGB(255 * vehtype1, 255 * vehtype1, 255 * vehtype1)
                                'Non-Blocking follower  ->  White   ->  RGB(255, 255, 255) -> RGB(255 * vehtype2, 255 * vehtype2, 255 * vehtype2)


                                If (foundblockingBoxMin = 1) Then
                                    If (blockingBoxVehiclesEB(minblockingBoxIndex).AttValue("Speed") < 15) Then
                                        Dim foundMax = 0
                                        Dim maxCoord As Double
                                        Dim maxIndex As Integer

                                        'all vehicles in the decisionbox that change lane while the value to block is 0 
                                        'then they get assigned to orange color and remain non-blocking type
                                        For t = 1 To decisionBoxVehiclesEB.Count
                                            If (decisionBoxVehiclesEB(t).AttValue("LANECHANGE") <> 0 And likelihood(value) = 0) Then
                                                decisionBoxVehiclesEB(t).AttValue("Type") = 2
                                                If (decisionBoxVehiclesEB(t).AttValue("Type") = 2) Then
                                                    decisionBoxVehiclesEB(t).AttValue("color") = RGB(255, 165, 0)
                                                End If
                                            End If

                                            'all vehicles in the decisionbox that change lane while the value to block greater than 0 
                                            'then they get assigned to purple color and become blocking type as they show aggresive tendency
                                            If (decisionBoxVehiclesEB(t).AttValue("LANECHANGE") <> 0 And likelihood(value) <> 0) Then
                                                decisionBoxVehiclesEB(t).AttValue("Type") = 1
                                                If (decisionBoxVehiclesEB(t).AttValue("Type") = 1) Then
                                                    decisionBoxVehiclesEB(t).AttValue("color") = RGB(255, 0, 255)
                                                End If
                                            End If
                                        Next

                                        'All the vehicles in the decisioBox are compared to find ID and coordinate of vehicle leading
                                        'in the decision box for each lane
                                        For p = 1 To decisionBoxVehiclesEB.Count
                                            If (decisionBoxVehiclesEB(p).AttValue("Lane") = n) Then
                                                'Console.WriteLine(decisionBoxVehiclesEB(p).AttValue("POINT").X)
                                                If (foundMax = 0) Then
                                                    maxCoord = decisionBoxVehiclesEB(p).AttValue("POINT").X
                                                    maxIndex = p
                                                    foundMax = 1
                                                ElseIf (decisionBoxVehiclesEB(p).AttValue("POINT").X > maxCoord) Then
                                                    maxCoord = decisionBoxVehiclesEB(p).AttValue("POINT").X
                                                    maxIndex = p
                                                End If
                                            End If
                                        Next

                                        'if the vehicle identified as leading in the decision box is unassigned to blocking or non-blocking type
                                        'then it is assigned to blocking purple or non-blocking orange depending on the likelihood value 
                                        'assigned for the run if
                                        If foundMax = 1 Then
                                            If (decisionBoxVehiclesEB(maxIndex).AttValue("Type") = 40) Then
                                                decisionBoxVehiclesEB(maxIndex).AttValue("Type") = flipcoin(likelihood(value)) + 1

                                                If (decisionBoxVehiclesEB(maxIndex).AttValue("Type") = 1) Then
                                                    decisionBoxVehiclesEB(maxIndex).AttValue("color") = RGB(255, 0, 255)
                                                End If

                                                If (decisionBoxVehiclesEB(maxIndex).AttValue("Type") = 2) Then
                                                    decisionBoxVehiclesEB(maxIndex).AttValue("color") = RGB(255, 165, 0)
                                                End If
                                            End If

                                            'all the follower vehicles in the decisionbox are assigned to the same type as the lead vehicle 
                                            'and appear as white-non blocking or black-blocking
                                            For k = 1 To decisionBoxVehiclesEB.Count
                                                If (decisionBoxVehiclesEB(k).AttValue("Lane") = n) Then
                                                    Dim vehicle = decisionBoxVehiclesEB(k)
                                                    If Not vehicle Is Nothing Then
                                                        If (k <> maxIndex) Then
                                                            If (decisionBoxVehiclesEB(k).AttValue("Type") = 40) And (decisionBoxVehiclesEB(maxIndex).AttValue("Type") = 1) Then
                                                                decisionBoxVehiclesEB(k).AttValue("Type") = 1
                                                                Dim vehtype1 As Integer = decisionBoxVehiclesEB(k).AttValue("Type") - 1
                                                                decisionBoxVehiclesEB(k).AttValue("color") = RGB(255 * vehtype1, 255 * vehtype1, 255 * vehtype1)

                                                            ElseIf (decisionBoxVehiclesEB(k).AttValue("Type") = 40) And (decisionBoxVehiclesEB(maxIndex).AttValue("Type") = 2) Then
                                                                decisionBoxVehiclesEB(k).AttValue("Type") = 2
                                                                Dim vehtype2 As Integer = decisionBoxVehiclesEB(k).AttValue("Type") - 1
                                                                decisionBoxVehiclesEB(k).AttValue("color") = RGB(255 * vehtype2, 255 * vehtype2, 255 * vehtype2)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next k
                                        End If
                                    End If
                                    ' End If
                                End If
                            Next
                        End If
                    Next
                    vissim.Exit()

                    Thread.Sleep(5000)

                    'Move output file - Raw Data Collection(.mer), Delay Raw (.vlr), Delay Compiled(.vlz) and Error files(.err)
                    My.Computer.FileSystem.CopyFile("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.mer", "C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Results\limDemand\singlefile\Seed" & randseed & "\Vol" & Volume & "\dc\value" & value & ".mer", True)
                    My.Computer.FileSystem.CopyFile("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.vlr", "C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Results\limDemand\singlefile\Seed" & randseed & "\Vol" & Volume & "\delay\raw\value" & value & ".vlr", True)
                    My.Computer.FileSystem.CopyFile("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.vlz", "C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Results\limDemand\singlefile\Seed" & randseed & "\Vol" & Volume & "\delay\comp\value" & value & ".vlz", True)
                    My.Computer.FileSystem.CopyFile("C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Files\DBTB_network_5thfeb2016.err", "C:\Users\asaroj3\Desktop\GRA\Spring 2016\Research\Vissim\Files\Expt5\Results\limDemand\singlefile\Seed" & randseed & "\Vol" & Volume & "\error\value" & value & ".err", True)
                Next
            Next
        Next
    End Sub

End Module
