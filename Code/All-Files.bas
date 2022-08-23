Sub macro_geral()

origem = "C:\SEU_CAMINHO"

arquivo = Dir(origem)

Do Until arquivo = ""
         
         Set wks = Workbooks.Open(pasta & origem)
         
        
        'SEU CODIGO'/'MACRO'



         wks.Close
 
arquivo = Dir()

Loop

End Sub


