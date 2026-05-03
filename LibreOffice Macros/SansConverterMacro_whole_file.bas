Sub BalaramToUnicode
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("ñ","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ç","ù","à","Ä","É","Ü","Å","È","Ì","Ï","Ö","Ò","Ë","Ç","Ñ","Ù")
out() = Array("ṣ","ā","ī","ū","ṛ","ṝ","ḷ","ṅ","ñ","ṭ","ḍ","ṇ","ś","ḥ","ṁ","Ā","Ī","Ū","Ṛ","Ṝ","Ṅ","Ñ","Ṭ","Ḍ","Ṇ","Ś","Ṣ","Ḥ")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
End Sub

Sub UnicodeToBalaram
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("ā","ī","ū","ṛ","ṝ","ḷ","ṅ","ñ","ṭ","ḍ","ṇ","ś","ḥ","ṁ","ṃ","Ā","Ī","Ū","Ṛ","Ṝ","Ṅ","Ñ","Ṭ","Ḍ","Ṇ","Ś","Ḥ","ṣ")
out() = Array("ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ç","ù","à","à","Ä","É","Ü","Å","È","Ì","Ï","Ö","Ò","Ë","Ç","Ù","ñ")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
End Sub

Sub BalaramToRussianCyrillic
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("ñ","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ç","ù","à","Ä","É","Ü","Å","È","Ì","Ï","Ö","Ò","Ë","Ç","Ñ","Ù","a","b","c","d","e","g","h","i","j","k","l","m","n","o","p","r","s","t","u","v","y","A","B","C","D","E","G","H","I","J","K","L","M","N","O","P","R","S","T","U","V","Y")
out() = Array("ш","а̄","ӣ","ӯ","р̣","р̣̄","л̣","н̇","н̃","т̣","д̣","н̣","ш́","х̣","м̇","А̄","Ӣ","Ӯ","Р̣","Р̣̄","Н̇","Н̃","Т̣","Д̣","Н̣","Ш́","Ш","Х̣","а","б","ч","д","е","г","х","и","дж","к","л","м","н","о","п","р","с","т","у","в","й","А","Б","Ч","Д","Е","Г","Х","И","Дж","К","Л","М","Н","О","П","Р","С","Т","У","В","Й")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
oReplace.SearchRegularExpression = True
oReplace.SearchString = "^е"
oReplace.ReplaceString = "э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchString = "^Е"
oReplace.ReplaceString = "Э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchRegularExpression = False
oReplace.SearchString = " е"
oReplace.ReplaceString = " э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchString = " Е"
oReplace.ReplaceString = " Э"
ThisComponent.ReplaceAll(oReplace)
End Sub

Sub BalaramToUkrainianCyrillic
Dim inp() As String, out() As String
Dim n As Long
Dim i As Long
Dim asp() As String
Dim oReplace As Object
inp() = Array("ñ","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ç","ù","à","Ä","É","Ü","Å","È","Ì","Ï","Ö","Ò","Ë","Ç","Ñ","Ù","a","b","c","d","e","g","h","i","j","k","l","m","n","o","p","r","s","t","u","v","y","A","B","C","D","E","G","H","I","J","K","L","M","N","O","P","R","S","T","U","V","Y")
out() = Array("ш","а̄","ī","ӯ","р̣","р̣̄","л̣","н̇","н̃","т̣","д̣","н̣","ш́","х̣","м̇","Ā","Ī","Ӯ","Р̣","Р̣̄","Н̇","Н̃","Т̣","Д̣","Н̣","Ш́","Ш","Х̣","а","б","ч","д","е","ґ","х","і","дж","к","л","м","н","о","п","р","с","т","у","в","й","А","Б","Ч","Д","Е","Ґ","Х","І","Дж","К","Л","М","Н","О","П","Р","С","Т","У","В","Й")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
asp() = Array("к","г","ч","ж","т","д","п","б")
For i = 0 To UBound(asp())
    oReplace.SearchString = asp(i) & "х"
    oReplace.ReplaceString = asp(i) & "г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = UCase(asp(i)) & "х"
    oReplace.ReplaceString = UCase(asp(i)) & "г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = asp(i) & "Х"
    oReplace.ReplaceString = asp(i) & "Г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = UCase(asp(i)) & "Х"
    oReplace.ReplaceString = UCase(asp(i)) & "Г"
    ThisComponent.ReplaceAll(oReplace)
Next i
End Sub

Sub RussianCyrillicToBalaram
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("дж","Дж","ш́","а̄","ӣ","ӯ","р̣","р̣̄","л̣","н̇","н̃","т̣","д̣","н̣","х̣","м̇","а","б","ч","д","э","е","г","х","и","к","л","м","н","о","п","р","с","т","у","в","й","ш")
out() = Array("j","J","ç","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ù","à","a","b","c","d","e","e","g","h","i","k","l","m","n","o","p","r","s","t","u","v","y","ñ")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
End Sub

Sub UkrainianCyrillicToBalaram
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("дж","Дж","ш́","а̄","ī","ӯ","р̣","р̣̄","л̣","н̇","н̃","т̣","д̣","н̣","х̣","м̇","кг","ґг","чг","жг","тг","дг","пг","бг","Кг","Ґг","Чг","Жг","Тг","Дг","Пг","Бг","а","б","ч","д","е","ґ","г","х","і","к","л","м","н","о","п","р","с","т","у","в","й","ш")
out() = Array("j","J","ç","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ù","à","kh","gh","ch","jh","th","dh","ph","bh","Kh","Gh","Ch","Jh","Th","Dh","Ph","Bh","a","b","c","d","e","g","h","h","i","k","l","m","n","o","p","r","s","t","u","v","y","ñ")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
End Sub

Sub HarvardKyotoToRussianCyrillic
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("RR","lR","A","I","U","R","G","J","T","D","N","z","S","H","M","a","b","c","d","e","g","h","i","j","k","l","m","n","o","p","r","s","t","u","v","y")
out() = Array("р̣̄","л̣","а̄","ӣ","ӯ","р̣","н̇","н̃","т̣","д̣","н̣","ш́","ш","х̣","м̇","а","б","ч","д","е","г","х","и","дж","к","л","м","н","о","п","р","с","т","у","в","й")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
oReplace.SearchRegularExpression = True
oReplace.SearchString = "^е"
oReplace.ReplaceString = "э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchString = "^Е"
oReplace.ReplaceString = "Э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchRegularExpression = False
oReplace.SearchString = " е"
oReplace.ReplaceString = " э"
ThisComponent.ReplaceAll(oReplace)
oReplace.SearchString = " Е"
oReplace.ReplaceString = " Э"
ThisComponent.ReplaceAll(oReplace)
End Sub

Sub HarvardKyotoToUkrainianCyrillic
Dim inp() As String, out() As String
Dim n As Long
Dim i As Long
Dim asp() As String
Dim oReplace As Object
inp() = Array("RR","lR","A","I","U","R","G","J","T","D","N","z","S","H","M","a","b","c","d","e","g","h","i","j","k","l","m","n","o","p","r","s","t","u","v","y")
out() = Array("р̣̄","л̣","а̄","ī","ӯ","р̣","н̇","н̃","т̣","д̣","н̣","ш́","ш","х̣","м̇","а","б","ч","д","е","ґ","х","і","дж","к","л","м","н","о","п","р","с","т","у","в","й")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
asp() = Array("к","г","ч","ж","т","д","п","б")
For i = 0 To UBound(asp())
    oReplace.SearchString = asp(i) & "х"
    oReplace.ReplaceString = asp(i) & "г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = UCase(asp(i)) & "х"
    oReplace.ReplaceString = UCase(asp(i)) & "г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = asp(i) & "Х"
    oReplace.ReplaceString = asp(i) & "Г"
    ThisComponent.ReplaceAll(oReplace)
    oReplace.SearchString = UCase(asp(i)) & "Х"
    oReplace.ReplaceString = UCase(asp(i)) & "Г"
    ThisComponent.ReplaceAll(oReplace)
Next i
End Sub

Sub BalaramToNoDiacritics
Dim inp() As String, out() As String
Dim n As Long
Dim oReplace As Object
inp() = Array("ñ","ä","é","ü","å","è","ÿ","ì","ï","ö","ò","ë","ç","ù","à","Ä","É","Ü","Å","È","Ì","Ï","Ö","Ò","Ë","Ç","Ñ","Ù")
out() = Array("s","a","i","u","r","r","l","n","n","t","d","n","s","h","m","A","I","U","R","R","N","N","T","D","N","S","S","H")
oReplace = ThisComponent.createReplaceDescriptor()
oReplace.SearchCaseSensitive = True
oReplace.SearchRegularExpression = False
For n = 0 To UBound(inp())
    oReplace.SearchString = inp(n)
    oReplace.ReplaceString = out(n)
    ThisComponent.ReplaceAll(oReplace)
Next n
End Sub
