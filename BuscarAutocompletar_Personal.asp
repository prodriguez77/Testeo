ESTOYT MODIFICANDO TODOS PARA QUE HAYA CONFLICOT
LOS AMO
Dim Escritura	
    ' Instancia los objetos de servidor
	'
	
	Set Ficha = Server.CreateObject("PrySIGE.Ficha")
	Ficha.CadenaConexion = Session("ConnectionString")
	
    Escritura=request.QueryString("term") & "<|>" & request.QueryString("Sede") & "<|>" & request.QueryString("Nivel") 

    'response.write Escritura
    'response.end

    if Ficha.Metodo_Generico_SP_Ejecuta("EV_AUTOCOMPLETAR_EMPLEADOS_SP",CStr(Escritura),"",true) = false then
            Response.redirect("../../Pantalla_De_Error.asp?Error=" & Ficha.Error)       
            Response.End  
    end if


if not Ficha.EOF then
    Ficha.MoveFirst
    'rsElementTem.Filter="Fullname like '*" & **request.QueryString("term")** & "*'"

    do while not Ficha.EOF
        if strArr<>"" then strArr=strArr & ","
        strArr=strArr & "{"&"""value"":""" & Capitalize(Ficha.fields("DSAYN"))  & """, ""id"":" & Ficha.fields("CDLegajo") & "}" 

        Ficha.MoveNext
    loop
end if

Response.Write "[" & strArr & "]"

Session.CodePage = 1252

%>