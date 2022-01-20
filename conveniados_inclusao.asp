<!DOCTYPE html>
<%@LANGUAGE = "VBSCRIPT" CODEPAGE="65001"%>
<!--#Include file="conexao.inc" -->
<html>
	<head>
		<meta charset="utf-8">
			<title>Controle de Processos</title>
			
</head>
	</head>
<body>
	<%
								
		'Recuperando dados do formulário preenchidp em conveniados_inclusao_index.asp
		nome                    = Request.Form("nome")
		cpf   				   	= Request.Form("cpf")
		convenio                = Request.Form("convenio")
		usuario                 = Request.Form("usuario")
		senha                   = Request.Form("senha")
		senha2                  = Request.Form("senha2")
			
		'Limpando lixos dos dados e prevenindo ataques
		'1 Removendo os espações em branco
		'2 Removendo as aspas simples
		
		nome    			    = UCASE(REPLACE(TRIM(nome), "'",""))
		cpf      				= REPLACE(TRIM(cpf),"'","")
		convenio 				= REPLACE(TRIM(convenio),"'","")
		usuario  				= LCASE(REPLACE(TRIM(usuario),"'",""))
		senha   				= REPLACE(TRIM(senha),"'","")
		senha2   				= REPLACE(TRIM(senha2),"'","")
			
		'validando o formulário
			'Verificando se o campo convênio possui apenas números
		
		if isNumeric(convenio) = False Then
			ValidaConvenio = "erro"
			OcorreuErro = "sim"
		 
		end if
			
		'Validando o campo senha
		'Verificando se a senha digitada possui no mínimo 6 e no máximo 10 caracteres
		if Len(senha) <6 OR Len(senha2)<6 OR Len(senha)>10 OR Len(senha2)>10 Then
			ValidaTamanhoSenha = "erro"
			OcorreuErro = "sim"
		end if
		
		'Verificando se as senhas digitadas são iguais
		if senha <> senha2 Then
			ValidaSenhasIguais = "erro"
			OcorreuErro = "sim"
		end if
		
		'Localizado usuário no banco para inibir duplicados
		StrConveniado = "SELECT Nome_Conveniado,Usuario, CPF_Conveniado FROM Conveniados WHERE Nome_Conveniado ='" & nome &"' OR Usuario ='" & usuario   &"'OR CPF_Conveniado = '"& cpf & "'" 
		Set rsSQL = conexao.Execute(StrConveniado)
			
						
		if not rsSQL.EOF Then
			ValidaUsuariosDiferentes = "erro"
			OcorreuErro = "sim"
							
		else
			if (ValidaConvenio <> "erro") and (ValidaTamanhoSenha <> "erro") and (ValidaSenhasIguais <> "erro") Then
				Response.Write "Não ocorreram erros no formulário" & "<br>"
				Set add_action = Server.CreateObject ("ADODB.Recordset")
                                             
                add_action.Open "Conveniados", conexao,3,3
                                             
                add_action.AddNew
                                                             
					add_action("Nome_Conveniado")= nome
					add_action("CPF_Conveniado") = cpf
					add_action("Convenio") = convenio
					add_action("Usuario") = usuario
					add_action("Senha") = senha
                                                               
				add_action.Update
                 
				add_action.Close
				conexao.Close
				
				Set add_action = Nothing
				Set conexao = Nothing
			end if
		end if
                
	%>
	<%
		if ValidaConvenio = "erro" Then
			Response.Write "<font style = 'color:red;'> O campo Convenio deve conter apenas números!" &" </font><br>"
		End If
		
		if ValidaTamanhoSenha = "erro" Then
			Response.Write "<font style='color:red;'> Sua senha deve ter no mínino 6 e no máximo 10 caracteres!" & " </font><br>"
		End if
		
		if ValidaSenhasIguais = "erro" Then
			Response.Write "<font style='color:red;'> Senhas não conferem"  &" </font><br>"
		end if
		
		if ValidaUsuariosDiferentes= "erro" Then
			Response.Write "<font style='color:red';> Existe um conveniado com o nome:" & nome &"</font><br>"
			Response.Write "<font style='color:red';> Ou existe um usuário:" & usuario&"</font><br>"
			Response.Write "<font style='color:red';> Ou existe um CPF:" & cpf &"</font><br>"
		end if
						
				
		if OcorreuErro = "sim" Then
	%>			
				 
	
				<div style="text-align: left; position: absolute; left:50%;top:20%;margin-left: -110px; margin-top: -40px">
				<p align="center"> Conveniados [<small>Inclusão</small>]</p>
					<form action="conveniados_inclusao.asp" method="post">
					<p align="left">
						<label for="nome">Nome completo:</label><br>
						<input onkeyup="this.value.replace(/[çÇáÁàÀèéÈÉíìÌÍòóÒÓùúÚÙñÑ~@&]/g,'')"
							type="texte name="
							nome"
							id="nome"
							required="requerid"
							placeholder="Nome Completo"
							style="text-transform: uppercase;"
							minlength="10"
							size="40"
							type="text"
							name="nome"
							id="nome"
							placeholder="Nome Completo"
							size="40"
							autofocus value = "<% =nome %>">
					</p>
					<p>
						<label for="cpf">CPF:</label><br>
						<input type="text"
							name="cpf"
							id="cpf"
							required="required"
							placeholder="Apenas Números"
							pattern="\d{11}"
							minlength="11"
							maxlength="11"
							autofocus value = "<% =cpf %>">
					</p>
					<p>
						<label for="convenio">Número do Convênio:</label><br>
							<input type="text"
								name="convenio"
								id="convenio"
								autocomplete="on"
								required="required"
								maxlength="10"
								pattern="[0-9]+$"
								placeholder="número do convênio"
								autofocus value = "<% =convenio %>">
					</p>
					<p>
						<label for="usuario">Usuário:</label><br>
							<input onkeyup="this.value=this.value.replace(/[' ']/g,'')"
								style="text-transform:lowercase;"
								required="requerid"
								input
								type="text"
								name="usuario"
								id="usuario"
								placeholder="Nome de Usuário"
								autofocus value = "<% =usuario %>">
					</p>
					<p>
						<label for="senha">Senha:</label><br>
						<input type="password"
							name="senha"
							id="senha"
							maxlength="10">
					</p>
					<p>
						<label for="senha2">Confirme a senha:</label><br>
							<input type="password"
								name="senha2"
								id="senha2"
								maxlength="10">
					</p>
					
					<input type="submit"name="submit"value="Incluir">
					<input type="reset"name="reset"value="Limpar">
					
					</form>
				<div>
	<%
		End If
	%>
	
</body>
</html>
