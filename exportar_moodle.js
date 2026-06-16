function ExportarMoodle() {
    let app = Application;
    let wb = app.ActiveWorkbook;
    let wsOrigem = wb.Sheets("Etapa 01. Preencher informações");
    
    let ano = String(new Date().getFullYear());
    let sigla   = String(wsOrigem.Cells(4, 1).Value2 || "");
    let unidade = String(wsOrigem.Cells(4, 3).Value2 || "");
    let curso   = String(wsOrigem.Cells(4, 7).Value2 || "");
    let nomeArq = `${sigla} - ${unidade} - ${curso} - ${ano}`;
    
    let pasta = "\\\\60.21.20.10\\SST - Saude e Seguranca$\\05. Suporte de sistemas\\04.06. Cadastro de usuários\\01.08.01. Histórico\\Moodle\\";
    let caminhoCSV  = `${pasta}CSV\\CSV - ${ano}\\${nomeArq}.csv`;
    let caminhoPDF  = `${pasta}PDF\\PDF - ${ano}\\${nomeArq}.pdf`;
    let caminhoXLSX = `${pasta}XLSX\\XLSX - ${ano}\\${nomeArq}.xlsx`;
    
    let ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "D").End(xlUp).Row;
    
    if (ultimaLinha < 4) {
        MsgBox("Nenhum dado encontrado.", 48, "Aviso");
        return;
    }
 


// --- 1. CSV ---
try {
    
    let wbCSV = app.Workbooks.Add();
    let wsCSV = wbCSV.Sheets(1);
    
   
    let cabecalhos = ["username", "firstname", "lastname", "email", "password", "institution", "department", "course1", "group1"];
    for (let c = 0; c < cabecalhos.length; c++) {
        wsCSV.Cells(1, c + 1).Value2 = cabecalhos[c];
    }
    
    
    for (let i = 4; i <= ultimaLinha; i++) {
        let row = i - 2; 
        
        let siglaUnid = String(wsOrigem.Cells(i, 2).Value2 || "").replace(/"/g, "");
        let nomeUnid  = String(wsOrigem.Cells(i, 3).Value2 || "").replace(/"/g, "");
        let nomeCompl = String(wsOrigem.Cells(i, 4).Value2 || "").replace(/"/g, "");
        let cpf       = String(wsOrigem.Cells(i, 5).Value2 || "").replace(/"/g, "").replace(/\./g, "").replace(/-/g, "");
        let email     = String(wsOrigem.Cells(i, 6).Value2 || "").replace(/"/g, "");
        let curso     = String(wsOrigem.Cells(i, 7).Value2 || "").replace(/"/g, "");
        let lastname  = "- " + siglaUnid + "/" + nomeUnid;
        let group1    = siglaUnid + "/" + nomeUnid;

    
        wsCSV.Cells(row, 1).Value2 = cpf;
        wsCSV.Cells(row, 2).Value2 = nomeCompl;
        wsCSV.Cells(row, 3).Value2 = lastname;
        wsCSV.Cells(row, 4).Value2 = email;
        wsCSV.Cells(row, 5).Value2 = "teste123";
        wsCSV.Cells(row, 6).Value2 = nomeUnid;
        wsCSV.Cells(row, 7).Value2 = siglaUnid;
        wsCSV.Cells(row, 8).Value2 = curso;
        wsCSV.Cells(row, 9).Value2 = group1;
    }
    
    app.DisplayAlerts = false;
    wbCSV.SaveAs(caminhoCSV, 6); 
    wbCSV.Close(false);
    app.DisplayAlerts = true;

} catch(err) {
    MsgBox("Erro CSV: " + err.message, 48, "Erro");
}
 // ---  PDF ---
    try {
        let wsPDF = wb.Sheets("Etapa 02. imprimir PDF");
        wsPDF.PrintOut(1, 1, 1, false, "", true, false, caminhoPDF);
    } catch(err) {
        MsgBox("Erro PDF: " + err.message, 48, "Erro");
    }
 
    // ---  XLSX ---
    try {
        let wbNovo = app.Workbooks.Add();
        let wsNovo = wbNovo.Sheets(1);
        wsNovo.Name = "Usuarios";
        let cab2 = ["sigla_associacao","sigla_unidade","nome_unidade","nome_completo","cpf","email","curso","senha"];
        for (let c = 0; c < cab2.length; c++) {
            wsNovo.Cells(1, c + 1).Value2 = cab2[c];
        }
        for (let i = 4; i <= ultimaLinha; i++) {
            let row = i - 2;
            for (let col = 1; col <= 7; col++) {
                wsNovo.Cells(row, col).Value2 = String(wsOrigem.Cells(i, col).Value2 || "").replace(/"/g, "");
            }
            wsNovo.Cells(row, 8).Value2 = "teste123";
        }
        app.DisplayAlerts = false;
        wbNovo.SaveAs(caminhoXLSX, 51);
        wbNovo.Close(false);
        app.DisplayAlerts = true;
    } catch(err) {
        MsgBox("Erro XLSX: " + err.message, 48, "Erro");
    }
 
    MsgBox(`Exportacao concluida!\n\nArquivos salvos em:\n${pasta}`, 64, "Moodle");
}
 function CommandButton1_Click()
{
    ExportarMoodle();
}
