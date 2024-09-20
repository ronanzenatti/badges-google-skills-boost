import axios from "axios";
import * as cheerio from "cheerio";
import readline from "readline";
import xlsx from "xlsx";
import { join } from "path";
import { promises as fs } from "fs";
import linkify from "linkifyjs";

var nomeArquivo = "";
var dadosPlanilha;
var base = 1;

var alunos = [];
var alunosInvalidos = [];

var totais = {
  totalAlunos: 0,
  perfisInvalidos: 0,
  perfisValidos: 0,
  perfisPrivados: 0,
  zerados: 0,
  gccf: 0,
  genAI: 0,
};

var skillsBadgesGCCF = [
  "Implement Load Balancing on Compute Engine",
  "Set Up an App Dev Environment on Google Cloud",
  "Build a Secure Google Cloud Network",
  "Prepare Data for ML APIs on Google Cloud",
];

var skillsBadgesGenAI = [
  "Prompt Design in Vertex AI",
  "Develop GenAI Apps with Gemini and Streamlit",
  "Inspect Rich Documents with Gemini Multimodality and Multimodal RAG",
];

// Cria uma interface para leitura do nome da planilha
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// INICIO do Programa
console.log(" ===== Levantamento de Badges do Google Skills Boost =====");

// Faz a pergunta do nome do arquivo para o usuário
rl.question("Qual o nome do arquivo? \n", (nome) => {
  nomeArquivo = nome;

  rl.question("BASE: 1 - CPS e 0 - FAT? \n", (tipoArquivo) => {
    base = tipoArquivo;
    // Fecha a interface
    rl.close();

    lerArquivo();
  });
});

function teste() {
  console.log(`Arquivo: ${nomeArquivo} - Base: ${base}`);
}

// Le o arquivo do Excel
async function lerArquivo() {
  console.log(" ===== Lendo arquivo do Excel =====");
  // Caminho completo do arquivo
  const filePath = join("input", nomeArquivo);

  // Verifica se o arquivo existe
  await fs.access(filePath);

  // Lê o arquivo XLSX
  const workbook = xlsx.readFile(filePath);

  // Obtém a planilha 'respostas'
  const nomePlanilha = "PERFIS";
  if (!workbook.SheetNames.includes(nomePlanilha)) {
    throw new Error(`A planilha "${nomePlanilha}" não existe no arquivo.`);
  }

  const worksheet = workbook.Sheets[nomePlanilha];
  console.log(" ===== Lendo planilha no arquivo do Excel =====");
  // Converte a planilha em JSON
  dadosPlanilha = xlsx.utils.sheet_to_json(worksheet);

  montaAlunos();
}

function montaAlunos() {
  console.log(" ===== Montando versão final dos alunos =====");
  // Percorre as linhas da tabela e exibe os valores das colunas específicas
  console.log(`Total de alunos: ${dadosPlanilha.length}`);
  dadosPlanilha.forEach((row, index) => {
    totais.totalAlunos++;

    if (base == 1) {
      console.log("=> Processando CPS");
      let { DATA, INSTITUCIONAL, CPF, PLATAFORMA, PERFIL } = row;
      console.log(`Linha ${index + 1}:`);
      console.log(` INSTITUCIONAL: ${INSTITUCIONAL} -  PERFIL: ${PERFIL}`);

      PERFIL = verificaURL(PERFIL);

      if (PERFIL) {
        alunos.push({
          cpf: CPF,
          institucional: INSTITUCIONAL,
          perfil: PERFIL,
          badges: [],
          gccf: false,
          genAI: false,
          faltando: "",
        });
        totais.perfisValidos++;
      } else {
        alunosInvalidos.push({ DATA, INSTITUCIONAL, CPF, PLATAFORMA, PERFIL });
        alunos.perfisInvalidos++;
      }
    } else if (base == 0) {
      console.log("=> Processando FAT");
      let { DATA, PLATAFORMA, PERFIL } = row;
      console.log(`Linha ${index + 1}:`);
      console.log(` EMAIL: ${PLATAFORMA} -  PERFIL: ${PERFIL}`);

      PERFIL = verificaURL(PERFIL);

      if (PERFIL) {
        alunos.push({
          email: PLATAFORMA,
          perfil: PERFIL,
          badges: [],
          gccf: false,
          genAI: false,
          faltando: "",
        });
        totais.perfisValidos++;
      } else {
        alunosInvalidos.push({ DATA, PLATAFORMA, PERFIL });
        alunos.perfisInvalidos++;
      }
    }
  });
  buscaBadges();
}

function verificaURL(perfil) {
  const link = linkify.find(perfil, "url");
  if (link.length > 0 && link[0].isLink) {
    let url = link[0].href;
    if (
      url.startsWith("https://www.cloudskillsboost.google/public_profiles/")
    ) {
      // Remove ?locale=pt_BR se existir no final da URL
      url = url.replace(/\?locale=pt_BR$/, "");
      return url;
    } else {
      return null;
    }
  } else {
    return null;
  }
}

async function buscaBadges() {
  console.log(" ===== Buscando Badges dos alunos com Perfil válido =====");
  for (const [index, row] of alunos.entries()) {
    try {
      const { data } = await axios.get(row.perfil);
      const $ = cheerio.load(data);

      // Pega o nome na plataforma
      alunos[index].nomeSkillsBost = $(".ql-display-small")
        .text()
        .trim()
        .toLocaleUpperCase();
      console.log(` [ Processando ] => ${alunos[index].nomeSkillsBost}`);

      // Verifica se ganhou alguma badge
      if ($(".ql-body-large.l-mtxl").length > 0) {
        alunos[index].badges = 0;
        alunos[index].gccf = "NÃO FINALIZADO";
        alunos[index].genAI = "NÃO FINALIZADO";
      } else {
        // Adiciona o total de Badges
        alunos[index].badges = $(".profile-badge").length;

        const badgesTemp = [];
        // Pega o nome de todas as badges e guarda temporáriamente.
        $(".profile-badge").each((index, item) => {
          badgesTemp.push($(item).find(".ql-title-medium").text().trim());
        });

        //  Valida se recebeu todas as do GCCF
        const badgesFaltandoGCCF = skillsBadgesGCCF.filter(
          (badge) =>
            !badgesTemp.some((tempBadge) => tempBadge.startsWith(badge))
        );
        alunos[index].gccf =
          badgesFaltandoGCCF.length === 0
            ? "CONCLUÍDO!"
            : badgesFaltandoGCCF.length < skillsBadgesGCCF.length
            ? `FALTAM ${badgesFaltandoGCCF.length} BADGES!`
            : "NÃO INICIADO!";

        alunos[index].faltando += "CLOUD: " + badgesFaltandoGCCF.join(", ");

        //  Valida se recebeu todas as do GenAI
        const badgesFaltandoGenAI = skillsBadgesGenAI.filter(
          (badge) =>
            !badgesTemp.some((tempBadge) => tempBadge.startsWith(badge))
        );
        alunos[index].genAI =
          badgesFaltandoGenAI.length === 0
            ? "CONCLUÍDO!"
            : badgesFaltandoGenAI.length < skillsBadgesGenAI.length
            ? `FALTAM ${badgesFaltandoGenAI.length} BADGES!`
            : "NÃO INICIADO!";

        alunos[index].faltando += " | GenAI: " + badgesFaltandoGenAI.join(", ");
      }
    } catch (error) {
      console.error(`Erro ao buscar dados para o aluno ${index}:`, error);
    }
  }

  toJSON();
}

function toJSON() {
  fs.writeFile(
    `output/alunos_${nomeArquivo.replace(".xlsx", "")}.json`,
    JSON.stringify(alunos),
    function (err) {
      if (err) {
        console.log(err);
      } else {
        console.log("JSON => Arquivo de [ alunos ] salvo!");
      }
    }
  );

  fs.writeFile(
    `output/alunos_invalidos_${nomeArquivo.replace(".xlsx", "")}.json`,
    JSON.stringify(alunosInvalidos),
    function (err) {
      if (err) {
        console.log(err);
      } else {
        console.log("JSON => Arquivo de [ alunos_invalidos ] salvo!");
      }
    }
  );
}
