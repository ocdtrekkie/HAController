Module modRandom
    Function BasicRandomInteger(ByVal intMax As Integer)
        ' Returns random integer between 1 and intMax.
        Return CInt(Int((intMax * Rnd()) + 1))
    End Function

    Function TrueRandomInteger(ByVal intMax As Integer)
        ' Experimental random.org coin flip request, does not yet work
        ' If I keep this, the code needs to fail gracefully to a local random flip
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
        ' API Key needs to be a setting
        Dim randomJSON As String = "{""jsonrpc"":""2.0"",""method"":""generateIntegers"",""params"":{""apiKey"":""" + My.Settings.Random_RandomOrgAPIKey + """,""n"":1,""min"":1,""max"":2,""replacement"":true,""base"":10},""id"":2000}"
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

        Return 0
    End Function
End Module
