import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import { 
  Document, Packer, Paragraph, Table, TableRow, TableCell, 
  WidthType, AlignmentType, TextRun, ImageRun 
} from 'docx';

const App = () => {
  const [responsavel, setResponsavel] = useState('');
  const [inicio, setInicio] = useState(1);
  const [fim, setFim] = useState(300);
  const [loading, setLoading] = useState(false);

  // Helper para normalizar nome do arquivo
  const normalizarNome = (nome) => {
    return nome
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-zA-Z0-9_]/g, "_")
      .toLowerCase();
  };

  // Função que transforma a imagem em Buffer (necessário para a Web)
  const getLogoBuffer = async () => {
    try {
      const response = await fetch('./logo-1.png'); // A imagem deve estar na pasta /public
      if (!response.ok) return null;
      return await response.arrayBuffer();
    } catch (e) {
      console.error("Erro ao carregar logo:", e);
      return null;
    }
  };

  const gerarWord = async () => {
    setLoading(true);
    const logoBuffer = await getLogoBuffer();
    const numeros = Array.from({ length: fim - inicio + 1 }, (_, i) => i + inicio);

    // Estrutura do Documento (Sua lógica original adaptada)
    const doc = new Document({
      sections: [{
        properties: {
          page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } }
        },
        children: [
          ...(logoBuffer ? [
            new Paragraph({
              children: [new ImageRun({ data: logoBuffer, transformation: { width: 120, height: 120 } })],
              alignment: AlignmentType.CENTER,
              spacing: { after: 300 }
            })
          ] : []),
          new Paragraph({
            children: [new TextRun({ text: "O Namorados da orgia está realizando uma rifa, valor R$ 2,00.", size: 22 })],
            spacing: { after: 100 }
          }),
          new Paragraph({
            children: [new TextRun({ text: `Responsável: ${responsavel}`, bold: true, size: 22 })],
            spacing: { after: 300 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: ["Número", "Nome", "Contato"].map(text => 
                  new TableCell({
                    children: [new Paragraph({ text, alignment: AlignmentType.CENTER })],
                    shading: { fill: "f0f0f0" }
                  })
                )
              }),
              ...numeros.map(num => new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ text: num.toString(), alignment: AlignmentType.CENTER })] }),
                  new TableCell({ children: [new Paragraph("")] }),
                  new TableCell({ children: [new Paragraph("")] })
                ]
              }))
            ]
          })
        ]
      }]
    });

    // Gerar e disparar download
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${normalizarNome(responsavel)}_${inicio}_a_${fim}.docx`);
    setLoading(false);
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h1 style={styles.title}>Gerador de Rifas</h1>
        <p style={styles.subtitle}>Namorados da Orgia</p>

        <div style={styles.formGroup}>
          <label>Nome do Responsável:</label>
          <input 
            type="text" 
            value={responsavel} 
            onChange={(e) => setResponsavel(e.target.value)}
            style={styles.input}
            placeholder="Ex: João Silva"
          />
        </div>

        <div style={styles.row}>
          <div style={styles.formGroup}>
            <label>Início:</label>
            <input 
              type="number" 
              value={inicio} 
              onChange={(e) => setInicio(Number(e.target.value))}
              style={styles.input}
            />
          </div>
          <div style={styles.formGroup}>
            <label>Fim:</label>
            <input 
              type="number" 
              value={fim} 
              onChange={(e) => setFim(Number(e.target.value))}
              style={styles.input}
            />
          </div>
        </div>

        <button 
          onClick={gerarWord} 
          disabled={!responsavel || loading}
          style={loading ? styles.buttonDisabled : styles.button}
        >
          {loading ? 'Gerando...' : 'Gerar Documento .docx'}
        </button>
      </div>
    </div>
  );
};

// CSS-in-JS básico para o design
const styles = {
  container: { display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: '100vh', backgroundColor: '#f4f7f6', fontFamily: 'Arial' },
  card: { backgroundColor: 'white', padding: '2rem', borderRadius: '12px', boxShadow: '0 4px 6px rgba(0,0,0,0.1)', width: '100%', maxWidth: '400px' },
  title: { textAlign: 'center', color: '#333', marginBottom: '0.5rem' },
  subtitle: { textAlign: 'center', color: '#666', marginBottom: '2rem' },
  formGroup: { marginBottom: '1rem', display: 'flex', flexDirection: 'column' },
  row: { display: 'flex', gap: '1rem' },
  input: { padding: '0.8rem', borderRadius: '6px', border: '1px solid #ddd', marginTop: '0.5rem' },
  button: { width: '100%', padding: '1rem', backgroundColor: '#007bff', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '1rem', marginTop: '1rem' },
  buttonDisabled: { width: '100%', padding: '1rem', backgroundColor: '#ccc', color: 'white', border: 'none', borderRadius: '6px', cursor: 'not-allowed', marginTop: '1rem' }
};

export default App;