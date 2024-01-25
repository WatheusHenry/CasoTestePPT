document.addEventListener("DOMContentLoaded", (event) => {
  document
    .getElementById("gerarSlideBtn")
    .addEventListener("click", gerarSlide);
});

function gerarSlide() {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();

  const rows = [
    ["Nome do Projeto", document.getElementById("input2").value],
    ["Sprint", document.getElementById("input3").value],
    ["História", document.getElementById("input4").value],
    ["Versão Testada", document.getElementById("input5").value],
    ["Status", document.getElementById("input6").value],
    ["Observação", document.getElementById("input7").value],
  ];

  const tableOptions = {
    x: 1,
    y: 2.3,
    w: 10,
    rowH: 0.5,
    colW: [3.0, 5.0],
    fontSize: 14,
    border: { type: "solid", pt: 1, color: "000000" },
    align: "center",
    fill: "F7F7F7",
  };

  slide.background = { color: "336699" };
  slide.addTable(rows, tableOptions);
  slide.addTable([["Caso de Teste", document.getElementById("input1").value]], {
    ...tableOptions,
    fontWeight: "bold",
    fill: "A9D08E",
    rowH: 1,
    y: 0.5,
  });

  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" }; // Definindo uma cor de fundo para o segundo slide

  // Função para adicionar imagem e texto ao slide
  function addImageToSlide(imageInputId, textInputId, xPosition) {
    return new Promise((resolve, reject) => {
      const textContent = document.getElementById(textInputId).value;
      const imageInput = document.getElementById(imageInputId);

      if (imageInput.files && imageInput.files[0]) {
        const file = imageInput.files[0];
        const reader = new FileReader();

        reader.onload = function (e) {
          const imageOptions = {
            data: e.target.result,
            x: xPosition,
            y: 1,
            w: 3,
            h: 4.5,
          };
          slide2.addImage(imageOptions);

          slide2.addText(textContent, {
            x: xPosition,
            y: 0,
            w: 3,
            h: 1,
            fontSize: 14,
            color: "000000",
          });

          resolve();
        };

        reader.onerror = reject;

        reader.readAsDataURL(file);
      } else {
        resolve(); // Resolve imediatamente se não houver arquivo
      }
    });
  }
  const nomeArquivo = document.getElementById("nomeArquivo").value || "Slide";

  // Executando todas as promessas de adicionar imagem
  Promise.all([
    addImageToSlide("imageInput", "inputTextImage", 0.3),
    addImageToSlide("imageInput2", "inputTextImage2", 3.5),
    addImageToSlide("imageInput3", "inputTextImage3", 6.7),
  ])
    .then(() => {
      pptx.writeFile(nomeArquivo + ".pptx");
    })
    .catch((error) => {
      console.error("Erro ao carregar imagens", error);
    });

  // document.getElementById("input1").value = "";
  // document.getElementById("input2").value = "";
  // document.getElementById("input3").value = "";
  // document.getElementById("input4").value = "";
  // document.getElementById("input5").value = "";
  // document.getElementById("input6").value = "";
  // document.getElementById("input7").value = "";
  // document.getElementById("inputTextImage").value = "";
  // document.getElementById("imageInput").value = "";
  // document.getElementById("inputTextImage2").value = "";
  // document.getElementById("imageInput2").value = "";
  // document.getElementById("inputTextImage3").value = "";
  // document.getElementById("imageInput3").value = "";
}
