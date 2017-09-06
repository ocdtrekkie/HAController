﻿Imports System.Xml

Module modMapQuest
    Sub Disable()
        My.Application.Log.WriteEntry("Unloading MapQuest module")
        Unload()
        My.Settings.MapQuest_Enable = False
        My.Application.Log.WriteEntry("MapQuest module is disabled")
    End Sub

    Sub Enable()
        My.Settings.MapQuest_Enable = True
        My.Application.Log.WriteEntry("MapQuest module is enabled")
        My.Application.Log.WriteEntry("Loading MapQuest module")
        Load()
    End Sub

    Function GetLocation(ByVal dblLatitude As Double, ByVal dblLongitude As Double) As String
        If My.Settings.MapQuest_Enable = True Then
            Dim LocationRequestString As String
            Dim LocationData As XmlDocument = New XmlDocument
            Dim LocationNode As XmlNode
            LocationRequestString = "http://www.mapquestapi.com/geocoding/v1/reverse?key=" + My.Settings.MapQuest_APIKey + "&location=" + CStr(dblLatitude) + "," + CStr(dblLongitude) + "&includeRoadMetadata=true&includeNearestIntersection=true&outFormat=xml"

            My.Application.Log.WriteEntry("Requesting MapQuest data")

            Try
                LocationData.Load(LocationRequestString)

                LocationNode = LocationData.SelectSingleNode("/response/results/result/locations/location/street")
                Dim strStreetName As String = LocationNode.InnerText
                My.Application.Log.WriteEntry(strStreetName)

                Return strStreetName
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)

                Return "Error getting location data"
            Catch XmlExcep As System.Xml.XmlException
                My.Application.Log.WriteException(XmlExcep)

                Return "Error parsing location data"
            End Try
        Else
            My.Application.Log.WriteEntry("MapQuest module is disabled")
            Return "Disabled"
        End If
    End Function

    Sub Load()
        If My.Settings.MapQuest_Enable = True Then
            If My.Settings.MapQuest_APIKey = "" Then
                My.Application.Log.WriteEntry("No MapQuest API key, asking for it")
                My.Settings.MapQuest_APIKey = InputBox("Enter MapQuest API Key. You can get an API key at https://developer.mapquest.com by signing up for a free account.", "MapQuest API")
            End If
        Else
            My.Application.Log.WriteEntry("MapQuest module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()

    End Sub
End Module
