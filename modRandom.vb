Module modRandom
    Function BasicRandomInteger(ByVal intMax As Integer)
        ' Returns random integer between 1 and intMax.
        Return CInt(Int((intMax * Rnd()) + 1))
    End Function

    Function RandomInteger(ByVal intMax As Integer)
        ' Returns random integer between 1 and intMax. Attempts to use Random.org if allowed and online.
        Dim intRandom As Integer = 0
        If My.Settings.Random_RandomOrgEnable = True And modGlobal.IsOnline = True Then
            My.Application.Log.WriteEntry("Contacting Random.org for random result.")
            intRandom = TrueRandomInteger(intMax)
            If intRandom = 0 Then
                My.Application.Log.WriteEntry("Random.org API use returned invalid input, using local random function instead.", TraceEventType.Warning)
                intRandom = BasicRandomInteger(intMax)
            End If
        Else
            My.Application.Log.WriteEntry("Random.org API use is disabled or the Internet is not connected, using local random function instead.")
            intRandom = BasicRandomInteger(intMax)
        End If
        Return intRandom
    End Function

    Function TrueRandomInteger(ByVal intMax As Integer)
        If My.Settings.Random_RandomOrgAPIKey = "00000000-0000-0000-0000-000000000000" Or My.Settings.Random_RandomOrgAPIKey = "" Then
            My.Application.Log.WriteEntry("No Random.Org API key, asking for it")
            My.Settings.Random_RandomOrgAPIKey = InputBox("Enter Random.Org API Key. You can get an API key at https://api.random.org/api-keys/beta by entering your email address.", "Random.org API")
        End If
        Dim randomAPIURL As String = "https://api.random.org/json-rpc/1/invoke"
        Dim randomRequest As System.Net.HttpWebRequest = System.Net.WebRequest.Create(randomAPIURL)
        randomRequest.Method = "POST"
        randomRequest.ContentType = "application/json"
        My.Application.Log.WriteEntry("Creating web request to " + randomAPIURL)
        Dim randomStream As System.IO.Stream = randomRequest.GetRequestStream()
        Dim randomJSON As String = "{""jsonrpc"":""2.0"",""method"":""generateIntegers"",""params"":{""apiKey"":""" + My.Settings.Random_RandomOrgAPIKey + """,""n"":1,""min"":1,""max"":" + CStr(intMax) + ",""replacement"":true,""base"":10},""id"":2000}"
        My.Application.Log.WriteEntry("Request sent: " + randomJSON)
        Dim requestBuffer As Byte() = System.Text.Encoding.UTF8.GetBytes(randomJSON)
        randomStream.Write(requestBuffer, 0, requestBuffer.Length)
        randomStream.Close()
        Dim randomResponse As System.Net.HttpWebResponse = randomRequest.GetResponse()
        Dim randomResponseStream As System.IO.Stream = randomResponse.GetResponseStream()
        Dim encode As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")
        Dim randomResponseRead As New System.IO.StreamReader(randomResponseStream, encode)
        Dim read(256) As Char
        Dim count As Integer = randomResponseRead.Read(read, 0, 256)
        Dim randomResponseJSON As String = ""
        While count > 0
            Dim str As New String(read, 0, count)
            randomResponseJSON = randomResponseJSON + str
            count = randomResponseRead.Read(read, 0, 256)
        End While
        randomResponseStream.Close()
        randomResponseRead.Close()
        My.Application.Log.WriteEntry("Response received: " + randomResponseJSON)
        Dim intRandom As Integer = CInt(randomResponseJSON.Substring(45, 1))
        If intRandom >= 1 And intRandom <= intMax Then
            Return intRandom
        Else
            My.Application.Log.WriteEntry("Invalid random result", TraceEventType.Warning)
            Return 0
        End If
    End Function
End Module
