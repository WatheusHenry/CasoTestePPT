document.addEventListener("DOMContentLoaded", (event) => {
  document
    .getElementById("gerarSlideBtn")
    .addEventListener("click", gerarSlide);
});

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("addInputBtn").addEventListener("click", addNewInput);
});

function addNewInput() {
  const container = document.getElementById("inputsContainer");

  // Criando a div que conterá os inputs
  const inputDiv = document.createElement("div");
  inputDiv.className = "formImages";

  // Criando o input de texto
  const textInput = document.createElement("input");
  textInput.type = "text";
  textInput.placeholder = "Insira o texto";
  textInput.className = "formImage";

  // Criando o input de arquivo (imagem)
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = "image/*";
  fileInput.className = "file-input";

  // Adicionando os inputs à div
  inputDiv.appendChild(textInput);
  inputDiv.appendChild(fileInput);

  // Adicionando a div ao container
  container.appendChild(inputDiv);
}

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
    x: 2,
    y: 1.5,
    w: "80%",
    rowH: 0.5,
    colW: [3.0, 3.0],
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

  const textContent = document.getElementById("inputTextImage").value; // Obtendo o texto do input
  const imageInput = document.getElementById("imageInput");

  if (imageInput.files && imageInput.files[0]) {
    const file = imageInput.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
      const imageOptions = {
        data: e.target.result,
        x: 0.3,
        y: 1,
        w: 3,
        h: 4,
      };
      slide2.addImage(imageOptions);

      slide2.addText(textContent, {
        x: 0.3,
        y: 0,
        w: 3,
        h: 1,
        fontSize: 14,
        color: "000000",
      });

      pptx.writeFile("CT0000.pptx");
    };

    reader.readAsDataURL(file);
  }
  const textContent2 = document.getElementById("inputTextImage2").value; // Obtendo o texto do input
  const imageInput2 = document.getElementById("imageInput2");
  if (imageInput2.files && imageInput2.files[0]) {
    const file = imageInput2.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
      const imageOptions = {
        data: e.target.result,
        x: 3.5,
        y: 1,
        w: 3,
        h: 4,
      };
      slide2.addImage(imageOptions);

      slide2.addText(textContent2, {
        x: 3.5,
        y: 0,
        w: 3,
        h: 1,
        fontSize: 14,
        color: "000000",
      });

      pptx.writeFile("CT0000.pptx");
    };

    reader.readAsDataURL(file);
  }
  const textContent3 = document.getElementById("inputTextImage3").value; // Obtendo o texto do input
  const imageInput3 = document.getElementById("imageInput3");
  if (imageInput3.files && imageInput3.files[0]) {
    const file = imageInput3.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
      const imageOptions = {
        data: e.target.result,
        x: 6.7,
        y: 1,
        w: 3,
        h: 4,
      };
      slide2.addImage(imageOptions);

      slide2.addText(textContent3, {
        x: 6.7,
        y: 0,
        w: 3,
        h: 1,
        fontSize: 14,
        color: "000000",
      });

      pptx.writeFile("CT0000.pptx");
    };

    reader.readAsDataURL(file);
  } else {
    // Caso nenhum arquivo seja selecionado, apenas salva o slide sem a imagem
    pptx.writeFile("CT0000.pptx");
  }

  document.getElementById("input1").value = "";
  document.getElementById("input2").value = "";
  document.getElementById("input3").value = "";
  document.getElementById("input4").value = "";
  document.getElementById("input5").value = "";
  document.getElementById("input6").value = "";
  document.getElementById("input7").value = "";
  document.getElementById("inputTextImage").value = "";
  document.getElementById("imageInput").value = "";
  document.getElementById("inputTextImage2").value = "";
  document.getElementById("imageInput2").value = "";
  document.getElementById("inputTextImage3").value = "";
  document.getElementById("imageInput3").value = "";
}
