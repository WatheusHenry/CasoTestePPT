document.addEventListener('DOMContentLoaded', (event) => {
  document.getElementById('gerarSlideBtn').addEventListener('click', gerarSlide);
});

function gerarSlide() {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();

  // Criar uma tabela com duas colunas e sete linhas
  const rows = [
      ['Caso de Teste', document.getElementById('input1').value],
      ['Nome do Projeto', document.getElementById('input2').value],
      ['Sprint', document.getElementById('input3').value],
      ['História', document.getElementById('input4').value],
      ['Versão Testada', document.getElementById('input5').value],
      ['Status', document.getElementById('input6').value],
      ['Observação', document.getElementById('input7').value]
  ];

  const tableOptions = {
      x: 2,
      y: 1,

      w: '80%',
      rowH: 0.5, // Altura da linha

      colW: [3.0, 3.0],
      fontSize: 14,
      border: { type: 'solid', pt: 1, color: '000000' }, // Bordas pretas
      align: 'center',
      fill: 'F7F7F7', // Cor de fundo das demais linhas
  };

  // Adicionar a tabela inteira
  slide.addTable(rows, tableOptions);
  // Personalizar a primeira linha separadamente
  slide.addTable([rows[0]], {
      ...tableOptions,
      fontWeight: 'bold',
      fill: 'A9D08E', // Cor de fundo da primeira linha
      rowH: 1, // Altura da linha
      y: 0.5 // Posição Y da tabela (mesma da tabela original)
  });

  // Personalizar o slide
  slide.background = {color:'336699'}; // Cor de fundo do slide



  pptx.writeFile('Slide_Com_Tabela_Estilizada.pptx');
}