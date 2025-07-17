require("dotenv").config(); // Carrega as variáveis de ambiente definidas no arquivo .env (ex: EMAIL_USER, EMAIL_PASS)
const express = require("express"); // Importa o framework Express para criar o servidor web
const multer = require("multer"); // Importa o Multer para lidar com uploads de arquivos via HTTP
const nodemailer = require("nodemailer"); // Importa o Nodemailer para envio de e-mails via SMTP
const cors = require("cors"); // Importa o CORS para permitir requisições cross-origin
const path = require("path"); // Importa módulos para manipulação de caminhos de arquivos e sistema de arquivos
const fs = require("fs"); // Importa o módulo 'fs' (file system) nativo do Node.js para manipulação de arquivos e diretórios no sistema operacional (leitura, escrita, criação, exclusão, etc.
const XLSX = require("xlsx"); // Importa a biblioteca XLSX para leitura e manipulação básica de arquivos Excel
const XlsxPopulate = require("xlsx-populate"); // Importa o XlsxPopulate para manipulação avançada e criação de arquivos Excel
const app = express(); // Cria uma instância do servidor Express
const port = 3000; // Define a porta onde o servidor ficará ouvindo
app.use(cors()); // Habilita CORS para aceitar requisições de outras origens (ex: front-end)
app.use(express.json()); // Habilita o Express para interpretar requisições com corpo JSON
app.use(express.static("public")); // Define a pasta "public" como estática para servir arquivos estáticos
const messageFilePath = path.join(__dirname, "message.json");
function processarMensagem(texto) {
  const dataAtual = new Date(); // Cria um objeto Date com a data e hora atual do sistema
  const meses = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ]; // Array com os nomes dos meses do ano em português, índice 0 = janeiro
  const mesAtual = dataAtual.getMonth(); // Obtém o mês atual como número (0 para janeiro, 1 para fevereiro, etc.)
  const mesAnterior = mesAtual === 0 ? 11 : mesAtual - 1; // Calcula o mês anterior: se for janeiro (0), retorna dezembro (11), senão subtrai 1
  const mesAtualNome = meses[mesAtual]; // Obtém o nome do mês atual a partir do array
  const mesAnteriorNome = meses[mesAnterior]; // Obtém o nome do mês anterior a partir do array
  const anoAtual = dataAtual.getFullYear(); // Obtém o ano atual com 4 dígitos (ex: 2025)
  return texto
    .replace(/<mes>/g, mesAtualNome) // Substitui todas as ocorrências de <mes> pelo nome do mês atual
    .replace(/<pmes>/g, mesAnteriorNome) // Substitui todas as ocorrências de <pmes> pelo nome do mês anterior
    .replace(/<ano>/g, anoAtual); // Substitui todas as ocorrências de <ano> pelo ano atual
}
// Rota GET para retornar a mensagem original (com placeholders <mes>, <ano>) para edição no modal
app.get("/mensagem-padrao", (req, res) => {
  // Lê o arquivo JSON que contém a mensagem padrão
  fs.readFile(messageFilePath, "utf8", (err, data) => {
    // Se ocorrer erro na leitura do arquivo, responde com status 500 e mensagem de erro
    if (err)
      return res.status(500).json({ erro: "Erro ao ler o arquivo JSON" });
    try {
      // Tenta converter o conteúdo JSON para objeto JavaScript
      const mensagem = JSON.parse(data); // Envia o conteúdo original da mensagem, sem substituir os placeholders
      res.json(mensagem);
    } catch {
      // Se o JSON estiver mal formatado, responde com erro 500
      res.status(500).json({ erro: "Erro ao analisar o JSON" });
    }
  });
});
// Rota GET para retornar a mensagem processada, com mês e ano atual substituídos
app.get("/mensagem-padrao/processada", (req, res) => {
  // Lê o arquivo JSON que contém a mensagem padrão
  fs.readFile(messageFilePath, "utf8", (err, data) => {
    // Se der erro na leitura, responde com status 500 e mensagem de erro
    if (err) {
      return res.status(500).json({ erro: "Erro ao ler o arquivo JSON" });
    }
    try {
      // Converte o JSON para objeto
      const mensagem = JSON.parse(data);
      // Usa a função processarMensagem para substituir os placeholders no assunto e no corpo
      const assuntoProcessado = processarMensagem(mensagem.assunto);
      const mensagemProcessada = processarMensagem(mensagem.mensagem);
      // Envia o assunto e mensagem já com os placeholders substituídos
      res.json({
        assunto: assuntoProcessado,
        mensagem: mensagemProcessada,
      });
    } catch (parseErr) {
      // Em caso de erro ao analisar o JSON, responde com erro 500
      res.status(500).json({ erro: "Erro ao analisar o JSON" });
    }
  });
});
// Rota POST para atualizar a mensagem padrão (assunto e mensagem)
// Recebe no corpo da requisição os dados atualizados e salva no arquivo JSON
app.post("/mensagem-padrao", (req, res) => {
  const novaMensagem = req.body; // Deve conter as propriedades "assunto" e "mensagem"
  // Salva o novo conteúdo no arquivo JSON, formatando com indentação para facilitar leitura
  fs.writeFile(
    messageFilePath,
    JSON.stringify(novaMensagem, null, 2),
    (err) => {
      // Se erro ao salvar o arquivo, retorna erro 500
      if (err) {
        return res.status(500).json({ erro: "Erro ao salvar o arquivo JSON" });
      }
      // Se sucesso, responde com objeto indicando sucesso
      res.json({ sucesso: true });
    }
  );
});
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms)); // Função que retorna uma Promise que resolve após um tempo (ms) especificado, útil para fazer "esperas" assíncronas
const storage = multer.diskStorage({
  // Configuração do storage para o multer, que define onde e como os arquivos enviados serão armazenados
  destination: (req, file, cb) => {
    // Define o destino (pasta) onde os arquivos serão salvos
    const now = new Date(); // Pega a data atual
    const ano = now.getFullYear().toString(); // Extrai o ano como string (ex: "2025")
    const mes = String(now.getMonth() + 1).padStart(2, "0"); // Extrai o mês atual, somando 1 pois getMonth() vai de 0 a 11, e formata para 2 dígitos (ex: "06")
    const dirAnoMes = path.join(__dirname, "uploads", ano, mes); // Cria o caminho da pasta onde o arquivo será salvo, por exemplo: "/caminho-do-projeto/uploads/2025/06"
    // Verifica se a pasta ainda não existe
    if (!fs.existsSync(dirAnoMes))
      //
      fs.mkdirSync(dirAnoMes, { recursive: true }); // Se não existir, cria a pasta recursivamente (cria todas as pastas necessárias na hierarquia)
    cb(null, dirAnoMes); // Callback do multer, passando null para indicar que não houve erro e informando o caminho da pasta
  },
  // Define o nome do arquivo que será salvo
  filename: (req, file, cb) => {
    const dataFormatada = new Date().toISOString().split("T")[0]; // Pega a data atual no formato ISO e extrai apenas a parte da data (ex: "2025-06-04")
    // Monta o nome do arquivo usando o prefixo "comissao_concreto_", a data formatada e a extensão original do arquivo
    // Exemplo: "comissao_concreto_2025-06-04.pdf"
    cb(
      null,
      `comissao_concreto_${dataFormatada}${path.extname(file.originalname)}`
    );
  },
});
const upload = multer({ storage }); // Cria o middleware Multer com a configuração de armazenamento definida acima

// Configura o serviço de envio de e-mail com Nodemailer usando Gmail e credenciais do .env
const transporter = nodemailer.createTransport({
  service: "gmail", // Usando Gmail para enviar
  auth: {
    user: process.env.EMAIL_USER, // Usuário e-mail (do .env)
    pass: process.env.EMAIL_PASS, // Senha do app ou conta (do .env)
  },
});
// Função para validar se todos os códigos da base estão na planilha enviada pelo usuário
function validarCodigosPlanilha(baseData, planilhaUsuarioData) {
  // Cria um conjunto (Set) com todos os códigos da base, garantindo que sejam strings sem espaços extras
  const codigosBase = new Set(baseData.map((item) => String(item.COD).trim())); // baseData é um array de objetos, cada um com a propriedade COD
  const codigosUsuario = new Set( // Cria um conjunto (Set) com todos os códigos da planilha enviada pelo usuário, também convertidos para string e limpos
    planilhaUsuarioData.map((item) => String(item.COD).trim())
  );
  let faltantes = []; // Inicializa um array para guardar os códigos que estão na base mas não foram encontrados na planilha do usuário
  // Para cada código que está na base, verifica se ele existe no conjunto de códigos do usuário
  codigosBase.forEach((codigo) => {
    // Se o código não existir na planilha do usuário, significa que está faltando
    if (!codigosUsuario.has(codigo)) {
      // Exibe no console uma mensagem avisando que o código está faltando na planilha
      console.log(
        `[VALIDAÇÃO] ⚠️ código ${codigo} não encontrado na planilha de vendedores comin_data.xlsx, verificar ocorrido !`
      );
      faltantes.push(codigo); // Adiciona o código faltante no array de faltantes
    }
  });
  // Se não houver nenhum código faltante (ou seja, todos foram encontrados)
  if (faltantes.length === 0) {
    // Exibe uma mensagem de validação bem-sucedida no console
    console.log(
      "[VALIDAÇÃO] ✅ Todos os códigos da base estão presentes na planilha enviada."
    );
  }
  // Retorna um objeto com a propriedade 'valido' como true
  // (obs: o retorno não considera se houve códigos faltantes ou não, poderia ser melhorado)
  return { valido: true };
}
// Rota POST para o envio de e-mails com a planilha
app.post("/enviar-email", upload.single("arquivo"), async (req, res) => {
  const inicio = Date.now(); // Marca o tempo de início da operação
  const logs = []; // Array para armazenar os logs capturados
  const originalLog = console.log; // Salva o log original para restaurar depois
  // Sobrescreve o console.log para também armazenar os logs no array
  console.log = function (...args) {
    const message = args
      .map((arg) => (typeof arg === "string" ? arg : JSON.stringify(arg)))
      .join(" ");
    logs.push(message); // Adiciona a mensagem ao array de logs
    originalLog.apply(console, args); // Mantém o comportamento original do log
  };
  // Verifica se algum arquivo foi enviado na requisição
  if (!req.file)
    return res.status(400).json({ error: "Nenhum arquivo enviado." });
  // Extrai os campos assunto e mensagem do corpo da requisição
  const { assunto, mensagem } = req.body;
  // Caminho do arquivo salvo pelo multer
  const filePath = req.file.path;
  // Logs de verificação de entrada
  console.log("\n\n");
  console.log(`[ASSUNTO] 📨: ${assunto || "Não informado"}`);
  console.log(" ");
  console.log(`[MENSAGEM] 📝: ${mensagem || "Não informado"}`);
  console.log(" ");
  console.log("[UPLOAD] ✅ Upload da planilha feito.");
  try {
    // Lê a planilha enviada
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets["apibase"]; // Busca a aba "apibase"
    // Verifica se a aba "apibase" existe
    if (!sheet)
      return res
        .status(400)
        .json({ error: 'A aba "apibase" não foi encontrada na planilha.' });
    // Converte os dados da planilha em JSON
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    // Lê a planilha base de vendedores (comin_data.xlsx)
    const vendedoresWorkbook = XLSX.readFile(
      path.join(__dirname, "comin_data.xlsx")
    );
    const abaConsulta = vendedoresWorkbook.Sheets["consulta_comin"];
    const vendedoresData = XLSX.utils.sheet_to_json(abaConsulta);
    // Valida os códigos da planilha enviada com os dados de cadastro
    validarCodigosPlanilha(vendedoresData, data);
    // Extrai os códigos únicos de vendedor
    const codigosUnicos = [...new Set(data.map((item) => item.COD))];
    const planilhasParaEnviar = []; // Array para armazenar os dados para envio
    const now = new Date(); // Data/hora atual
    const ano = now.getFullYear().toString(); // Ano atual
    const mes = String(now.getMonth() + 1).padStart(2, "0"); // Mês atual com 2 dígitos
    // Geração de planilhas por código
    for (const codigo of codigosUnicos) {
      // Filtra os dados por código
      const dadosVendedor = data
        .filter((item) => item.COD === codigo)
        .sort((a, b) =>
          (a["CL. DESC"] || "").localeCompare(b["CL. DESC"] || "", "pt-BR")
        );
      // Ignora se não houver dados
      if (dadosVendedor.length === 0) continue;
      // Busca as informações do vendedor no arquivo comin_data.xlsx
      const infoVendedor = vendedoresData.find((v) => v.COD === codigo);
      if (!infoVendedor) {
        console.log(
          `[AVISO] ⚠️ Código ${codigo} encontrado mas não possui cadastro.`
        );
        continue;
      }
      // Coleta os e-mails relacionados ao vendedor
      const emails = [
        infoVendedor.PRINCIPAL,
        infoVendedor.ALTERNATIVO,
        infoVendedor.REGIONAL,
        infoVendedor.CONTROLADORIA,
      ]
        .filter((e) => !!e) // Remove nulos/vazios
        .map((e) => e.trim().toLowerCase());
      // Verifica se há ao menos um e-mail
      if (emails.length === 0) {
        console.log(
          `[AVISO] ⚠️ Nenhum e-mail registrado no código ${codigo} no comin_data.xlsx.`
        );
        continue;
      }
      // Cria o diretório ano/mês para salvar os relatórios
      const dirAnoMes = path.join(__dirname, "relatorios", ano, mes);
      if (!fs.existsSync(dirAnoMes))
        fs.mkdirSync(dirAnoMes, { recursive: true });
      // Define o caminho do novo arquivo Excel
      const newFilePath = path.join(
        dirAnoMes,
        `comissao_${codigo}_${now.toISOString().split("T")[0]}.xlsx`
      );
      // Cria uma nova planilha em branco com XlsxPopulate
      const wb = await XlsxPopulate.fromBlankAsync();
      const sh = wb.sheet(0);
      sh.name(`Vendedor_${codigo}`);
      const headers = Object.keys(dadosVendedor[0]); // Cabeçalhos
      // Define os cabeçalhos com estilos
      headers.forEach((header, i) => {
        const cell = sh.cell(1, i + 1);
        cell.value(header).style({
          bold: true,
          fill: "4472C4",
          fontColor: "FFFFFF",
          horizontalAlignment: "center",
          verticalAlignment: "center",
          border: true,
        });
        sh.column(i + 1).width(Math.max(15, header.length + 5)); // Largura da coluna
      });
      // Preenche os dados na planilha
      // Loop pelos dados de cada vendedor para preencher a planilha
      dadosVendedor.forEach((linha, rowIndex) => {
        // Alterna a cor de fundo (efeito zebra): branco para pares, azul claro para ímpares
        const fillColor = rowIndex % 2 === 0 ? "FFFFFF" : "D9E1F2";
        // Loop pelas colunas (headers) para preencher cada célula
        headers.forEach((col, colIndex) => {
          // Define a célula atual com base na linha (offset +2 por causa do cabeçalho) e coluna
          const cell = sh.cell(rowIndex + 2, colIndex + 1);
          // Define o valor da célula com o conteúdo correspondente no objeto "linha"
          cell.value(linha[col]);
          // Aplica o estilo básico (cor de fundo e borda)
          cell.style({
            fill: fillColor,
            border: true,
          });
          // Se for uma coluna numérica que representa valores financeiros, formata como moeda
          if (
            ["TOTAL", "COMISSÃO", "VALOR", "REDUZIDA", "NORMAL"].includes(
              col.toUpperCase()
            )
          ) {
            // Aplica o formato de moeda e mantém borda e cor de fundo
            cell.style({
              fill: fillColor,
              border: true,
              numberFormat: '"R$" #,##0.00;[Red]-"R$" #,##0.00',
            });
            // Se o valor for negativo, aplica negrito para destacá-lo
            if (cell.value() < 0) {
              cell.style({ bold: true });
            }
          }
        });
      });
      // Após preencher os dados, ajusta a largura de cada coluna baseada na largura atual
      headers.forEach((_, colIndex) => {
        const col = sh.column(colIndex + 1); // Obtém a coluna pelo índice
        const larguraAtual = col.width(); // Lê a largura atual da coluna
        col.width(larguraAtual + 2); // Aumenta um pouco (ajuste sutil)
      });
      // Define largura fixa (explícita) para colunas extras
      sh.column("L").width(2); // L é usada apenas como separação visual (estreita)
      sh.column("M").width(10); // M será o cabeçalho "TOTAL"
      sh.column("N").width(14); // N conterá o subtotal com mais espaço
      // Cria o cabeçalho "TOTAL" na célula M1 com formatação destacada
      sh.cell("M1").value("TOTAL").style({
        bold: true, // Negrito
        fill: "4472C4", // Azul escuro (estilo cabeçalho)
        fontColor: "FFFFFF", // Texto branco
        horizontalAlignment: "center", // Alinhado horizontalmente ao centro
        verticalAlignment: "center", // Alinhado verticalmente ao centro
        border: true, // Borda ao redor
      });
      // Cria a célula N1 com fórmula de subtotal (respeita filtros), formatada como moeda
      const totalCell = sh.cell("N1").formula("SUBTOTAL(109,K:K)").style({
        fill: "FFFF00", // Amarelo (destaque visual)
        bold: true, // Negrito
        border: true, // Borda
        numberFormat: '"R$" #,##0.00;[Red]-"R$" #,##0.00', // Formato de moeda com negativo em vermelho
        horizontalAlignment: "center", // Centralizado na horizontal
        verticalAlignment: "center", // Centralizado na vertical
      });
      // Salva o arquivo gerado
      await wb.toFileAsync(newFilePath);
      console.log(`[RELATÓRIO] 📄 Planilha ${codigo} criada.`);
      // Adiciona à fila de envio
      planilhasParaEnviar.push({
        filePath: newFilePath,
        emails,
        codigo,
      });
    }
    console.log(
      `[RELATÓRIO] ✅ Todas as planilhas foram criadas. Iniciando envio dos e-mails...`
    );
    let emailsEnviados = 0;
    // Envia os e-mails com os relatórios anexados
    for (const item of planilhasParaEnviar) {
      const mailOptions = {
        from: process.env.EMAIL_USER,
        to: item.emails,
        subject: assunto || `Comissão - Vendedor ${item.codigo}`,
        text: mensagem || "Segue a planilha com a comissão.",
        attachments: [
          {
            filename: path.basename(item.filePath),
            path: item.filePath,
          },
        ],
      };
      try {
        // Envia o e-mail
        await transporter.sendMail(mailOptions);
        emailsEnviados++;
        console.log(
          `[ENVIO] ✉️ E-mail enviado para o vendedor ${
            item.codigo
          }: ${item.emails.join(", ")}`
        );
      } catch (err) {
        // Caso ocorra erro no envio de um e-mail
        console.error(
          `[ERRO] ❌ Erro ao enviar e-mail para ${item.codigo}:`,
          err
        );
      }
      await wait(1000); // Aguarda 1 segundo entre envios
    }
    // Finaliza o tempo de execução
    const finalizacao = Date.now();
    const duracaoSegundos = Math.floor((finalizacao - inicio) / 1000);
    const minutos = Math.floor(duracaoSegundos / 60);
    const segundos = duracaoSegundos % 60;
    // Formata o horário da finalização
    const horarioFormatado = new Date(finalizacao).toLocaleString("pt-BR", {
      timeZone: "America/Sao_Paulo",
    });
    // Logs finais
    console.log(" ");
    console.log(`[ENVIO] ✅ Todos os e-mails foram enviados com sucesso.`);
    console.log(`[EMAIL] 📬 Total de e-mails enviados: ${emailsEnviados}`);
    console.log(`[TEMPO] 🕒 Finalizado em: ${horarioFormatado}`);
    console.log(
      `[TEMPO] ⏱️ Duração do processo: ${minutos} minuto(s) e ${segundos} segundo(s)`
    );
    console.log(" ");
    // Salva o log em arquivo
    const logDir = path.join(__dirname, "log", ano, mes);
    if (!fs.existsSync(logDir)) fs.mkdirSync(logDir, { recursive: true });
    const dataBR = now
      .toLocaleDateString("pt-BR", { timeZone: "America/Sao_Paulo" })
      .replace(/\//g, "-");
    const horaBR = now
      .toLocaleTimeString("pt-BR", {
        timeZone: "America/Sao_Paulo",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      })
      .replace(/:/g, "-");
    const nomeArquivoLog = `log_${dataBR}_${horaBR}_COMIN.txt`;
    const logPath = path.join(logDir, nomeArquivoLog);
    fs.writeFileSync(logPath, logs.join("\n"), "utf-8");
    // Restaura o console.log original
    console.log = originalLog;
    console.log(`✅ Log salvo em: ${logPath}`);
    // Retorna resposta de sucesso
    return res.json({
      success: true,
      message: "Planilhas geradas e e-mails enviados com sucesso.",
      logPath,
    });
  } catch (error) {
    // Bloco de tratamento de erro geral
    console.log = originalLog; // Restaura console
    console.error("[ERRO] ❌ Erro geral:", error); // Loga o erro
    return res
      .status(500)
      .json({ error: "Erro interno ao processar o envio." }); // Retorna erro 500
  }
});
// Inicializa o servidor para escutar requisições na porta definida na variável port
app.listen(port, () => {
  console.log(`🚀 Servidor rodando na porta ${port}`);
});
// Desenvolvido por Eduardo Junqueira || contato: eduardojunqueira2005@gmail.com
