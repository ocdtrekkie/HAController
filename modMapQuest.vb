﻿Imports System.Xml

Module modMapQuest
    Function Disable() As String
        Unload()
        My.Settings.MapQuest_Enable = False
        My.Application.Log.WriteEntry("MapQuest module is disabled")
        Return "MapQuest module disabled"
    End Function

    Function Enable() As String
        My.Settings.MapQuest_Enable = True
        My.Application.Log.WriteEntry("MapQuest module is enabled")
        Load()
        Return "MapQuest module enabled"
    End Function

    Function GetDirections(ByVal strOrigin As String, ByVal strDestination As String) As String
        If My.Settings.MapQuest_Enable = True Then
            Dim DirectionsRequestString As String
            Dim DirectionsData As XmlDocument = New XmlDocument
            Dim DirectionsNodeList As XmlNodeList
            Dim DirectionsNode As XmlNode
            DirectionsRequestString = "http://www.mapquestapi.com/directions/v2/route?key=" + My.Settings.MapQuest_APIKey + "&from=" + System.Net.WebUtility.UrlEncode(strOrigin) + "&to=" + System.Net.WebUtility.UrlEncode(strDestination) + "&outFormat=xml"
            My.Application.Log.WriteEntry("Request string: " + DirectionsRequestString)

            My.Application.Log.WriteEntry("Requesting MapQuest data")

            Try
                DirectionsData.Load(DirectionsRequestString)

                DirectionsNodeList = DirectionsData.SelectNodes("/response/route/legs/leg/maneuvers/maneuver")
                modGPS.DirectionsListSize = DirectionsNodeList.Count
                ReDim modGPS.DirectionsNarrative(DirectionsListSize)
                ReDim modGPS.DirectionsLatitudeList(DirectionsListSize)
                ReDim modGPS.DirectionsLongitudeList(DirectionsListSize)

                For Each DirectionsNode In DirectionsNodeList
                    My.Application.Log.WriteEntry(DirectionsNode.Item("narrative").InnerText())
                    modGPS.DirectionsNarrative(CInt(DirectionsNode.Item("index").InnerText())) = DirectionsNode.Item("narrative").InnerText()
                    If DirectionsNode.Item("startPoint").HasChildNodes = True Then
                        ' The last maneuver has no "startPoint", so we only populate these fields if the startPoint exists.
                        modGPS.DirectionsLatitudeList(CInt(DirectionsNode.Item("index").InnerText())) = CDbl(DirectionsNode.Item("startPoint").Item("lat").InnerText())
                        modGPS.DirectionsLongitudeList(CInt(DirectionsNode.Item("index").InnerText())) = CDbl(DirectionsNode.Item("startPoint").Item("lng").InnerText())
                    End If
                    My.Application.Log.WriteEntry(modGPS.DirectionsLatitudeList(CInt(DirectionsNode.Item("index").InnerText())) & ", " & modGPS.DirectionsLongitudeList(CInt(DirectionsNode.Item("index").InnerText())), TraceEventType.Verbose)
                Next

                isNavigating = True
                Return modGPS.DirectionsNarrative(0)
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)

                Return "Error getting navigation data"
            Catch XmlExcep As System.Xml.XmlException
                My.Application.Log.WriteException(XmlExcep)

                Return "Error parsing navigation data"
            End Try
        Else
            My.Application.Log.WriteEntry("MapQuest module is disabled")
            Return "Disabled"
        End If
    End Function

    Function GetLocation(ByVal dblLatitude As Double, ByVal dblLongitude As Double) As String
        If My.Settings.MapQuest_Enable = True Then
            Dim LocationRequestString As String
            Dim LocationData As XmlDocument = New XmlDocument
            Dim LocationNode As XmlNode
            LocationRequestString = "http://www.mapquestapi.com/geocoding/v1/reverse?key=" + My.Settings.MapQuest_APIKey + "&location=" + CStr(dblLatitude) + "," + CStr(dblLongitude) + "&outFormat=xml"

            My.Application.Log.WriteEntry("Requesting MapQuest data")

            Try
                LocationData.Load(LocationRequestString)

                LocationNode = LocationData.SelectSingleNode("/response/results/result/locations/location/geocodeQuality")
                Dim strGeocodeQuality As String = LocationNode.InnerText

                If strGeocodeQuality = "ADDRESS" OrElse strGeocodeQuality = "STREET" Then
                    LocationNode = LocationData.SelectSingleNode("/response/results/result/locations/location/street")
                    Dim strStreetName As String = LocationNode.InnerText
                    Return strStreetName
                Else
                    My.Application.Log.WriteEntry("Geocode quality available: " + strGeocodeQuality)
                    Return "Poor quality result"
                End If
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

    Function Load() As String
        If My.Settings.MapQuest_Enable = True Then
            My.Application.Log.WriteEntry("Loading MapQuest module")
            If My.Settings.MapQuest_APIKey = "" Then
                My.Application.Log.WriteEntry("No MapQuest API key, asking for it")
                My.Settings.MapQuest_APIKey = InputBox("Enter MapQuest API Key. You can get an API key at https://developer.mapquest.com by signing up for a free account.", "MapQuest API")
            End If
            Return "MapQuest module loaded"
        Else
            My.Application.Log.WriteEntry("MapQuest module is disabled, module not loaded")
            Return "MapQuest module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading MapQuest module")
        Return "MapQuest module unloaded"
    End Function
End Module
