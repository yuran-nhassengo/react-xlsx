import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelReader = () => {
  const [data, setData] = useState<string[][]>([]);  // Armazenar os dados da planilha como array de arrays de strings
  const [fileName, setFileName] = useState<string>('');  // Armazenar o nome do arquivo carregado
  const [sheetNames, setSheetNames] = useState<string[]>([]);  // Armazenar os nomes das planilhas (abas)
  const [selectedSheet, setSelectedSheet] = useState<string>('');  // Armazenar a planilha selecionada
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null); // Armazenar o objeto do workbook

  // Função para ler e processar o arquivo Excel
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files ? event.target.files[0] : null;
    if (file) {
      setFileName(file.name);  // Definir o nome do arquivo

      const reader = new FileReader();

      reader.onload = (e) => {
        const binaryString = e.target?.result as string;
        const wb = XLSX.read(binaryString, { type: 'binary' });
        setWorkbook(wb); // Armazenar o workbook
        const sheetNames = wb.SheetNames;
        setSheetNames(sheetNames); // Atualizar as planilhas disponíveis

        if (selectedSheet) {
          loadSheetData(wb, selectedSheet); // Carregar dados da planilha selecionada, se já houver seleção
        }
      };

      reader.readAsBinaryString(file);
    }
  };

  // Função para carregar os dados de uma planilha selecionada
  const loadSheetData = (workbook: XLSX.WorkBook, sheetName: string) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    setData(jsonData as string[][]);
  };

  // Função para alterar a planilha selecionada
  const handleSheetChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const selectedSheetName = event.target.value;
    setSelectedSheet(selectedSheetName);
    if (workbook) {
      loadSheetData(workbook, selectedSheetName); // Carregar dados da planilha selecionada
    }
  };

  return (
    <div style={{ padding: '20px' }}>
      <h1>Carregar Arquivo Excel</h1>

      {/* Formulário para carregar o arquivo Excel */}
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        style={{ marginBottom: '20px' }}
      />
      {fileName && <p>Arquivo carregado: {fileName}</p>}

      {/* Se houver planilhas, exiba o seletor para escolher uma planilha */}
      {sheetNames.length > 0 && (
        <div>
          <label htmlFor="sheetSelector">Escolher Planilha:</label>
          <select
            id="sheetSelector"
            onChange={handleSheetChange}
            value={selectedSheet}
            style={{ margin: '10px' }}
          >
            <option value="">Selecione uma planilha</option>
            {sheetNames.map((sheetName, index) => (
              <option key={index} value={sheetName}>
                {sheetName}
              </option>
            ))}
          </select>
        </div>
      )}

      {/* Exibindo a tabela com os dados da planilha */}
      {data.length > 0 && (
        <div>
          <h2>Dados da Planilha:</h2>
          <table style={{ borderCollapse: 'collapse', width: '100%', border: '1px solid black' }}>
            <thead>
              <tr>
                {data[0].map((header: string, index: number) => (
                  <th key={index} style={{ padding: '8px', textAlign: 'left' }}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.slice(1).map((row: string[], index: number) => (
                <tr key={index}>
                  {row.map((cell: string, cellIndex: number) => (
                    <td key={cellIndex} style={{ padding: '8px', textAlign: 'left' }}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default ExcelReader;
