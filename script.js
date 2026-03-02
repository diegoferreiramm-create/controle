// 1. Defina a URL correta APENAS UMA VEZ no topo do arquivo
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxeyoKG99zETrrx6BdF7--w_-1cVe-S0tctxKOAfgFFQ3_as64oRqONoditWtXWsrRF/exec";

// --- VARIÁVEIS GLOBAIS DE CONTROLE ---
let modoEdicao = false;
let idSendoEditado = null;
let alunoEncontradoGlobal = null;

// --- FUNÇÃO AUXILIAR PARA CHAMADAS API ---
async function chamarAPI(params) {
  const query = new URLSearchParams(params).toString();
  const urlFinal = `${SCRIPT_URL}?${query}`;

  try {
    const response = await fetch(urlFinal);
    // Se o Google não responder 200 OK, ele pula para o catch
    if (!response.ok) throw new Error("Erro na rede");
    
    return await response.json(); 
  } catch (error) {
    console.error("Erro na API:", error);
    throw error; 
  }
}

function abrirTela(id){
  const telas = ["loginBox","menuBox","cadastrarBox","pesquisarBox","entregarBox","listasBox", "logBox", "recebimentoLoteBox", "corrigirBox"];   
  
  telas.forEach(t => { 
    const el = document.getElementById(t);
    if(el) el.style.display = "none"; 
  });

  const telaDestino = document.getElementById(id);
  if(telaDestino){
    telaDestino.style.display = "flex";
  }

  if(id === 'entregarBox') {
    document.getElementById("codigoCtr").value = "";
    document.getElementById("infoAlunoEntrega").style.display = "none";
    alunoEncontradoGlobal = null;
  }
  if(id === 'listasBox') carregarLista();
  if(id === 'logBox') carregarDadosLog();
  
  if(id !== 'cadastrarBox') {
    modoEdicao = false;
    idSendoEditado = null;
    const btnSalvar = document.getElementById("btnSalvar");
    if(btnSalvar) btnSalvar.innerText = "Salvar e Gerar Protocolo";
  }
}

async function entrar(){
  const login = document.getElementById("login").value.trim();
  const senha = document.getElementById("senha").value.trim();
  const msg = document.getElementById("msg");
  
  try {
    const res = await chamarAPI({ acao: "validarLogin", u: login, s: senha });
    if(res.sucesso){
      sessionStorage.setItem("usuario", JSON.stringify(res));
      mostrarMenu();
    } else { 
      msg.innerText="Login inválido"; 
    }
  } catch (err) {
    console.error("Erro no login:", err);
    msg.innerText = "Erro de conexão com o servidor.";
  }
}

const colunasDef = [
  { label: "ID", idx: 1 }, { label: "CPF", idx: 2 }, { label: "NOME", idx: 3 },
  { label: "NASC", idx: 4 }, { label: "MUNICIPIO", idx: 5 }, { label: "TEL", idx: 6 },
  { label: "VIA", idx: 7 }, { label: "PARCEIRO", idx: 8 }, 
  { label: "ATENDENTE", idx: 10 }, 
  { label: "Nº BOLETO", idx: 11 }, 
  { label: "STATUS", idx: 12 },    
  { label: "MOTIVO", idx: 13 },    
  { label: "DT ATU", idx: 14 },    
  { label: "CARTEIRA", idx: 15 },  
  { label: "LOTE", idx: 16 },      
  { label: "EDITAR", idx: 17 }     
];

const colunasParaMarcar = ["ID", "CPF", "NOME", "NASC", "MUNICIPIO", "TEL", "ATENDENTE", "Nº BOLETO", "EDITAR"];

function gerarChecksColunas() {
  const container = document.getElementById("containerChecks");
  if (!container) return;
  container.innerHTML = ""; 

  colunasDef.forEach((col) => {
    const label = document.createElement("label");
    label.style.cssText = "margin-right:10px; display:inline-flex; align-items:center; gap:5px; cursor:pointer; font-size:12px; white-space:nowrap;";
    const marcado = colunasParaMarcar.includes(col.label) ? "checked" : "";

    label.innerHTML = `<input type="checkbox" ${marcado} onchange="filtrarTabelaAvancado()" data-idx="${col.idx}"> ${col.label}`;
    container.appendChild(label);
  });
  
  setTimeout(filtrarTabelaAvancado, 200);
}

function filtrarTabelaAvancado() {
  const sessao = sessionStorage.getItem("usuario");
  if (!sessao) return;
  const user = JSON.parse(sessao);
  const isAdmin = (user.parceiro.toString() === "97");

  const fNome = document.getElementById("fNome").value.toUpperCase();
  const fStatus = document.getElementById("fStatus").value.toUpperCase();
  const fLote = document.getElementById("fLote") ? document.getElementById("fLote").value.trim() : "";
 
  const fParc = isAdmin ? document.getElementById("fParceiro").value.toUpperCase() : "";
  const fMun = isAdmin ? document.getElementById("fMun").value.toUpperCase() : "";
  const fAtend = isAdmin ? document.getElementById("fAtend").value.toUpperCase() : "";
  const fVia = (isAdmin && document.getElementById("fVia")) ? document.getElementById("fVia").value.toUpperCase() : "";

  const tabela = document.getElementById("tabelaListas");
  const tr = tabela.getElementsByTagName("tr");
  let contadorVisiveis = 0;

  for (let i = 1; i < tr.length; i++) {
    const td = tr[i].getElementsByTagName("td");
    if (!td[0]) continue;

    let mostrar = true;
    if (fNome && td[2] && td[2].innerText.toUpperCase().indexOf(fNome) === -1) mostrar = false;
    if (fStatus && td[11] && td[11].innerText.toUpperCase().trim() !== fStatus) mostrar = false;

    let txtLote = td[16] ? td[16].innerText.trim() : "";
    if (fLote && txtLote !== fLote) {
        if (td[15] && td[15].innerText.trim() === fLote) {
            mostrar = true; 
        } else {
            mostrar = false;
        }
    }
   
    if (isAdmin) {
      if (fVia && td[6] && td[6].innerText.toUpperCase().indexOf(fVia) === -1) mostrar = false;
      if (fParc && td[7] && td[7].innerText.toUpperCase().indexOf(fParc) === -1) mostrar = false;
      if (fMun && td[4] && td[4].innerText.toUpperCase().indexOf(fMun) === -1) mostrar = false;
      if (fAtend && td[9] && td[9].innerText.toUpperCase().indexOf(fAtend) === -1) mostrar = false;
    }

    tr[i].style.display = mostrar ? "" : "none";
    if (mostrar) contadorVisiveis++;
  }

  const elNumLinhas = document.getElementById("numLinhas");
  if (elNumLinhas) elNumLinhas.innerText = contadorVisiveis;

  const checks = document.querySelectorAll('#containerChecks input[type="checkbox"]');
  checks.forEach((input) => {
    const idx = input.getAttribute('data-idx');
    const visivel = input.checked;
    const colunas = tabela.querySelectorAll(`tr > *:nth-child(${idx})`);
    colunas.forEach(cel => {
      cel.style.display = visivel ? "" : "none";
    });
  });
}

function imprimirLista() {
  const conteudo = document.getElementById('areaImpressao').innerHTML;
  const telaPrint = window.open('', '_blank');
  
  telaPrint.document.write(`
    <html>
    <head>
      <title>Relatório</title>
      <style>
        @page { size: landscape; margin: 0.5cm; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #000; padding: 2px; font-size: 8px; }
        [style*="display: none"] { display: none !important; }
        th:last-child, td:last-child { display: none !important; }
      </style>
    </head>
    <body>
      ${conteudo}
      <script>window.onload = function() { window.print(); window.close(); };<\/script>
    </body>
    </html>
  `);
  telaPrint.document.close();
}

window.onload = gerarChecksColunas;

function atualizarRelogio() {
  const agora = new Date();
  const dia = String(agora.getDate()).padStart(2, '0');
  const mes = String(agora.getMonth() + 1).padStart(2, '0');
  const ano = agora.getFullYear();
  const horas = String(agora.getHours()).padStart(2, '0');
  const minutos = String(agora.getMinutes()).padStart(2, '0');
  const segundos = String(agora.getSeconds()).padStart(2, '0');
  
  const strDataHora = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
  const el = document.getElementById("dataHoraHud");
  if(el) el.innerText = strDataHora;
}
setInterval(atualizarRelogio, 1000);

const CPF = {
  limpar(valor){ return valor.replace(/\D/g,''); },
  formatar(valor){
    let v = this.limpar(valor);
    v = v.slice(0,11);
    v = v.replace(/(\d{3})(\d)/,'$1.$2');
    v = v.replace(/(\d{3})(\d)/,'$1.$2');
    v = v.replace(/(\d{3})(\d{1,2})$/,'$1-$2');
    return v;
  },
  validar(valor){
    let cpf = this.limpar(valor);
    if(cpf.length !== 11 || /^(\d)\1{10}$/.test(cpf)) return false;
    let soma=0, resto;
    for(let i=1;i<=9;i++) soma+=parseInt(cpf.substring(i-1,i))*(11-i);
    resto=(soma*10)%11; if(resto===10||resto===11) resto=0;
    if(resto!=parseInt(cpf.substring(9,10))) return false;
    soma=0;
    for(let i=1;i<=10;i++) soma+=parseInt(cpf.substring(i-1,i))*(12-i);
    resto=(soma*10)%11; if(resto===10||resto===11) resto=0;
    return resto==parseInt(cpf.substring(10,11));
  }
};

document.addEventListener('blur', async function(e) {
  if (e.target.id === 'cpf' && !modoEdicao) {
    const valorCpf = e.target.value;
    if (!CPF.validar(valorCpf)) return;
    document.getElementById("msgCPF").innerText = "Consultando base de dados...";
    
    try {
      const res = await chamarAPI({ acao: "buscarDadosNoBD", cpf: valorCpf });
      if (res && res.encontrado) {
        document.getElementById("msgCPF").innerText = "Dados recuperados!";
        document.getElementById("nome").value = res.nome || "";
        document.getElementById("municipio").value = res.municipio || "";
        document.getElementById("telefone").value = res.telefone || "";
        if (res.nascimento) {
          try {
            let dataStr = res.nascimento.toString();
            let dataFinal = "";
            if (dataStr.includes('/')) {
              const p = dataStr.split('/');
              dataFinal = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
            } else {
              let d = new Date(res.nascimento);
              if (!isNaN(d.getTime())) {
                let ano = d.getFullYear();
                let mes = String(d.getMonth() + 1).padStart(2, '0');
                let dia = String(d.getDate()).padStart(2, '0');
                dataFinal = `${ano}-${mes}-${dia}`;
              }
            }
            if (dataFinal) document.getElementById("nascimento").value = dataFinal;
          } catch(err) { console.error(err); }
        }
      } else {
        document.getElementById("msgCPF").innerText = "CPF não encontrado (Novo cadastro).";
      }
    } catch (err) {
      console.error(err);
      document.getElementById("msgCPF").innerText = "Erro ao consultar base.";
    }
  }
}, true);

function mostrarMenu(){
  const user = JSON.parse(sessionStorage.getItem("usuario"));
  if(!user) return;
  
  document.getElementById("infoUsuario").innerText = user.nome + " | " + user.parceiro;
  document.getElementById("hudUsuario").style.display="flex";

  if(user.nome === 'admin' || user.parceiro.toString() === "97") {
    const cardLog = document.getElementById("cardLog");
    if(cardLog) cardLog.style.display = "block";
  }

  abrirTela('menuBox');
}

async function carregarDadosLog() {
  const tbody = document.getElementById("corpoLogs");
  if(!tbody) return;
  tbody.innerHTML = "<tr><td colspan='5' style='text-align:center; padding:20px;'>Carregando Logs da Planilha...</td></tr>";

  try {
    const dados = await chamarAPI({ acao: "buscarLogsDaAbaLog" });
    tbody.innerHTML = "";
    if (!dados || dados.length === 0) {
      tbody.innerHTML = "<tr><td colspan='5' style='text-align:center; padding:20px;'>Aba LOG está vazia.</td></tr>";
      return;
    }
    dados.forEach(l => {
      tbody.innerHTML += `
        <tr style="border-bottom: 1px solid #334155;">
          <td style="padding:10px;">${l.data}</td>
          <td>${l.usuario}</td>
          <td style="color:#fbbf24;">${l.parceiro || '-'}</td>
          <td>${l.acao}</td>
          <td style="color:#22c55e;">${l.idRef}</td>
        </tr>`;
    });
  } catch (err) {
    tbody.innerHTML = "<tr><td colspan='5' style='text-align:center;'>Erro ao carregar logs.</td></tr>";
  }
}

function logout(){ sessionStorage.clear(); document.getElementById("hudUsuario").style.display="none"; abrirTela('loginBox'); }

function toggleTerceiro() {
  const isChecked = document.getElementById("checkTerceiro").checked;
  const container = document.getElementById("camposTerceiro");
  const inputs = container.querySelectorAll("input, select");
  container.style.opacity = isChecked ? "1" : "0.5";
  inputs.forEach(el => el.disabled = !isChecked);
}

function prepararEdicao(item) {
  modoEdicao = true;
  idSendoEditado = item.id;
  
  abrirTela('cadastrarBox');
  
  document.getElementById("cpf").value = item.cpf || "";
  document.getElementById("nome").value = item.nome || "";
  
  if(item.nasc) {
    const p = item.nasc.split('/');
    if(p.length === 3) {
      document.getElementById("nascimento").value = `${p[2]}-${p[1]}-${p[0]}`;
    }
  }
  
  document.getElementById("municipio").value = item.municipio || "";
  document.getElementById("telefone").value = item.tel || "";
  
  const campoBoleto = document.getElementById("codigoBoleto");
  if(campoBoleto) {
      campoBoleto.value = item.boleto || "";
  }
  
  document.getElementById("btnSalvar").innerText = "ATUALIZAR CADASTRO";
}

async function salvarCadastro(){
  const userStr = sessionStorage.getItem("usuario");
  if(!userStr) { alert("Sessão expirada. Faça login novamente."); return; }
  const user = JSON.parse(userStr);
  const cpf = document.getElementById("cpf").value;
  const nome = document.getElementById("nome").value;
  const nascRaw = document.getElementById("nascimento").value;
  const mun = document.getElementById("municipio").value;
  const tel = document.getElementById("telefone").value;
  
  const viaEl = document.querySelector('input[name="via"]:checked');
  const via = viaEl ? viaEl.value : "1ª VIA";
  
  const boleto = document.getElementById("codigoBoleto").value.trim();

  if(!cpf || !nome || !nascRaw || !boleto) { 
    alert("ERRO: CPF, Nome, Nascimento e Número do Boleto são obrigatórios!"); 
    return; 
  }

  const partes = nascRaw.split("-");
  const nascFormatado = `${partes[2]}/${partes[1]}/${partes[0]}`;
  const btn = document.getElementById("btnSalvar");
  btn.innerText = "Processando...";
  btn.disabled = true;

  // --- CONFIGURAÇÃO DA URL (COLOQUE A SUA URL AQUI) ---
  const urlWebApp = "SUA_URL_DO_WEB_APP_AQUI"; 

  try {
    if (modoEdicao) {
      // PREPARAÇÃO DOS PARÂMETROS PARA EDIÇÃO
      const params = new URLSearchParams({
        acao: "editarCadastroAppsScript",
        id: idSendoEditado,
        cpf: cpf,
        nome: nome,
        nasc: nascFormatado,
        municipio: mun,
        tel: tel,
        via: via,
        atendente: user.nome,
        parceiro: user.parceiro,
        boleto: boleto
      });

      const response = await fetch(`${urlWebApp}?${params.toString()}`);
      const res = await response.json();

      if(res.sucesso) {
        ["cpf", "nome", "nascimento", "municipio", "telefone", "codigoBoleto"].forEach(id => {
          const el = document.getElementById(id);
          if(el) el.value = "";
        });
        document.getElementById("msgCPF").innerText = "";
        abrirTela('listasBox');
      } else { alert("Erro ao editar: " + res.erro); }

    } else {
      // PREPARAÇÃO DOS PARÂMETROS PARA NOVO CADASTRO
      const res = await chamarAPI({
      acao: "salvarCadastroAppsScript",
      cpf: cpf,
      nome: nome,
      nasc: nascFormatado,
      municipio: mun,  // <-- Aqui: o .gs quer 'municipio', seu código tem 'mun'
      tel: tel,
      via: via,
      parceiro: user.parceiro,
      atendente: user.nome,
      boleto: boleto
    });

      const response = await fetch(`${urlWebApp}?${params.toString()}`);
      const res = await response.json();

      if(res.sucesso){ 
        imprimirProtocolo(res.id, cpf, nome, nascFormatado, mun, via, user.nome, user.parceiro, res.data, boleto);
        ["cpf", "nome", "nascimento", "municipio", "telefone", "codigoBoleto"].forEach(id => {
          const el = document.getElementById(id);
          if(el) el.value = "";
        });
        document.getElementById("msgCPF").innerText = "";
      } else { alert("Erro: " + res.erro); }
    }
  } catch (err) {
    console.error(err);
    alert("Erro de conexão com o servidor. Verifique a URL e a permissão 'Qualquer pessoa'.");
  } finally {
    btn.innerText = "Salvar e Gerar Protocolo";
    btn.disabled = false;
  }
}

} function salvarCadastro(){
  const userStr = sessionStorage.getItem("usuario");
  if(!userStr) { alert("Sessão expirada. Faça login novamente."); return; }
  const user = JSON.parse(userStr);
  const cpf = document.getElementById("cpf").value;
  const nome = document.getElementById("nome").value;
  const nascRaw = document.getElementById("nascimento").value;
  const mun = document.getElementById("municipio").value;
  const tel = document.getElementById("telefone").value;
  
  const viaEl = document.querySelector('input[name="via"]:checked');
  const via = viaEl ? viaEl.value : "1ª VIA";
  
  const boleto = document.getElementById("codigoBoleto").value.trim();

  if(!cpf || !nome || !nascRaw || !boleto) { 
    alert("ERRO: CPF, Nome, Nascimento e Número do Boleto são obrigatórios!"); 
    return; 
  }

  const partes = nascRaw.split("-");
  const nascFormatado = `${partes[2]}/${partes[1]}/${partes[0]}`;
  const btn = document.getElementById("btnSalvar");
  btn.innerText = "Processando...";
  btn.disabled = true;

  try {
    if (modoEdicao) {
      const res = await chamarAPI({
        acao: "editarCadastroAppsScript",
        id: idSendoEditado,
        cpf: cpf,
        nome: nome,
        nasc: nascFormatado,
        mun: mun,
        tel: tel,
        via: via,
        atendente: user.nome,
        parceiro: user.parceiro,
        boleto: boleto
      });

      if(res.sucesso) {
        ["cpf", "nome", "nascimento", "municipio", "telefone", "codigoBoleto"].forEach(id => {
          const el = document.getElementById(id);
          if(el) el.value = "";
        });
        document.getElementById("msgCPF").innerText = "";
        abrirTela('listasBox');
      } else { alert("Erro ao editar: " + res.erro); }

    } else {
      const res = await chamarAPI({
        acao: "salvarCadastroAppsScript",
        cpf: cpf,
        nome: nome,
        nasc: nascFormatado,
        mun: mun,
        tel: tel,
        via: via,
        parceiro: user.parceiro,
        atendente: user.nome,
        boleto: boleto
      });

      if(res.sucesso){ 
        imprimirProtocolo(res.id, cpf, nome, nascFormatado, mun, via, user.nome, user.parceiro, res.data, boleto);
        ["cpf", "nome", "nascimento", "municipio", "telefone", "codigoBoleto"].forEach(id => {
          const el = document.getElementById(id);
          if(el) el.value = "";
        });
        document.getElementById("msgCPF").innerText = "";
      } else { alert("Erro: " + res.erro); }
    }
  } catch (err) {
    alert("Erro de conexão com o servidor.");
  } finally {
    btn.innerText = "Salvar e Gerar Protocolo";
    btn.disabled = false;
  }
}

function imprimirProtocolo(id, cpf, nome, nascimento, municipio, via, atendente, parceiro, data, boleto) {
  const telaPrint = window.open('', '_blank');
  telaPrint.document.write(`
    <html>
    <head>
      <title>Protocolo CTR - ${id}</title>
      <style>
        @page { size: A5 landscape; margin: 0; }
        body { font-family: Arial, sans-serif; padding: 5mm; color: #000; }
        .ticket { border: 2px solid #000; padding: 10px; width: 195mm; height: 135mm; display: flex; flex-direction: column; box-sizing: border-box; }
        .header { text-align: center; border-bottom: 2px solid #000; margin-bottom: 10px; }
        .id-destaque { font-size: 24px; font-weight: bold; }
        .row { display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 14px; }
        .lgpd { font-size: 9px; font-style: italic; margin-top: 10px; border-top: 1px solid #ccc; padding-top: 5px; text-align: justify; }
        .rules { font-size: 10px; background: #f2f2f2; padding: 8px; border: 1px solid #000; margin-top: 8px; line-height: 1.3; }
        .footer { display: flex; justify-content: space-between; align-items: flex-end; margin-top: auto; padding-bottom: 5px; }
        b { text-transform: uppercase; }
      </style>
    </head>
    <body>
      <div class="ticket">
        <div class="header">
          <h2 style="margin:5px 0;">PROTOCOLO DE SOLICITAÇÃO</h2>
          <span class="id-destaque">Nº BOLETO: ${boleto}</span>
        </div>
        <div class="content">
          <div class="row"><span><b>NOME:</b> ${nome.toUpperCase()}</span><span><b>DATA:</b> ${data.split(' ')[0]}</span></div>
          <div class="row"><span><b>CPF:</b> ${cpf}</span><span><b>VIA:</b> ${via}</span></div>
          <div class="row"><span><b>MUNICÍPIO:</b> ${municipio.toUpperCase()}</span><span><b>ATENDENTE:</b> ${atendente}</span></div>
          <div class="lgpd">Não nos responsabilizamos por informações no formulário entregue que divergirem dos documentos anexos, conforme Art. 9º da Lei 13.709/2018 (LGPD). A veracidade é de responsabilidade do declarante.</div>
          <div class="rules">
            <strong>Procedimento para Entrega da Carteira Estudantil:</strong><br>
            • Aluno, mãe, pai, irmãos ou filhos: Apresentar o comprovante de solicitação original e um documento oficial com foto.<br>
            • (Em caso de perda ou extravio do comprovante, apresentar uma cópia do documento oficial com foto de quem for receber.)<br>
            • Tios, primos, demais parentes ou terceiros: Apresentar o comprovante de solicitação original e um documento oficial com foto de quem estiver recebendo, juntamente com uma cópia do documento oficial do aluno.<br><br>
            <strong>EM HIPÓTESE ALGUMA ENTREGAREMOS A TERCEIROS SEM O COMPROVANTE DE SOLICITAÇÃO ORIGINAL EM MÃOS.</strong>
          </div>
        </div>
        <div class="footer">
          <div style="border-top:1px solid #000; width:250px; text-align:center; font-size:12px; margin-top: 20px;">Assinatura do Requerente</div>
          <div style="font-size:12px;">Via do Aluno / ${parceiro} / CTR: ${id}</div>
        </div>
      </div>
      <script>window.print();<\/script>
    </body>
    </html>
  `);
  telaPrint.document.close();
}

async function salvarSenha() {
  const login = document.getElementById("usuarioTroca").value.trim();
  const atual = document.getElementById("senhaAtual").value.trim();
  const nova = document.getElementById("novaSenha").value.trim();
  const conf = document.getElementById("confSenha").value.trim();
  
  if(nova !== conf) { alert("A nova senha não coincide!"); return; }
  
  try {
    const res = await chamarAPI({ 
      acao: "trocarSenha", 
      login: login, 
      atual: atual, 
      nova: nova 
    });
    if(res.sucesso) { alert("Senha alterada!"); fecharSenha(); }
    else { alert("Erro ao alterar senha."); }
  } catch (err) {
    alert("Erro de conexão ao alterar senha.");
  }
}

async function executarBuscaGeral(tipo) {
  const user = JSON.parse(sessionStorage.getItem("usuario"));
  const valor = document.getElementById("valorPesquisa").value;
  if(!valor) return alert("Digite algo para pesquisar");
  
  const div = document.getElementById("resultadoPesquisa");
  div.innerHTML = "Pesquisando...";
  
  try {
    const res = await chamarAPI({ 
      acao: "pesquisarNoCadastroGeral", 
      valor: valor, 
      tipo: tipo, 
      parceiro: user.parceiro 
    });
    
    div.innerHTML = "";
    
    if(!res || res.length === 0) {
      div.innerHTML = "Nenhum registro encontrado ou sem permissão.";
      return;
    }
    
    res.forEach(item => {
      let dStat = "";
      for(let key in item) {
        if(key.toUpperCase().replace(/\s/g,'') === "DATASTATUS") dStat = item[key];
      }
      if(!dStat) dStat = item.dataStatus || item.data_status || item["DATA STATUS"] || "";

      div.innerHTML += `
        <div class="res-card">
          <b>NOME:</b> ${item.nome}<br>
          <b>CPF:</b> ${item.cpf} | <b>VIA:</b> ${item.via}<br>
          <b>MUNICÍPIO:</b> ${item.municipio} | <b>PARCEIRO:</b> ${item.parceiro}<br>
          <b>CARTEIRA:</b> ${item.numCarteira || 'N/A'}<br>
          <b>STATUS:</b> ${item.status || 'Pendente'} | <b>MOTIVO:</b> ${item.motivo || '-'}<br>
          <b>DATA STATUS:</b> ${dStat}<br>
          <b>ATENDENTE:</b> ${item.atendente}<br>
          <small>Última Atualização: ${item.dataAtu}</small>
        </div>
      `;
    });
  } catch (err) {
    div.innerHTML = "Erro ao processar pesquisa.";
  }
}

document.addEventListener('input', function (e) {
  const camposCpfObrigatorio = ['cpf', 'cpfTerceiro'];

  if (camposCpfObrigatorio.includes(e.target.id)) {
      e.target.value = CPF.formatar(e.target.value);
  }
  
  if(e.target.id === 'cpf') {
    const v = CPF.validar(e.target.value);
    const msg = document.getElementById("msgCPF");
    if(msg) {
      msg.innerText = v ? "CPF Válido" : "CPF Inválido";
      msg.className = v ? "valid" : "invalid";
    }
    const btnSalvar = document.getElementById("btnSalvar");
    if(btnSalvar) btnSalvar.disabled = !v;
  }

  if(e.target.id === 'codigoCtr') e.target.value = e.target.value.replace(/\D/g, "");
});

function abrirSenha(){ document.getElementById("modalSenha").style.display="flex"; }
function fecharSenha(){ document.getElementById("modalSenha").style.display="none"; }

async function carregarLista() {
  const user = JSON.parse(sessionStorage.getItem("usuario"));
  const isAdmin = (user.parceiro.toString() === "97");
  
  const fAdmin = document.getElementById("filtrosAdmin");
  if(fAdmin) {
    fAdmin.style.display = "flex";
    fAdmin.style.flexDirection = "row";
    fAdmin.style.flexWrap = "wrap";
    fAdmin.style.gap = "5px";
    fAdmin.style.alignItems = "center";
    fAdmin.style.padding = "10px";
    fAdmin.style.background = "#0f172a";
    fAdmin.style.borderRadius = "8px";
    fAdmin.style.marginBottom = "10px";
    
    if(!document.getElementById("fLote")){
       const inputLote = document.createElement("input");
       inputLote.id = "fLote";
       inputLote.placeholder = "LOTE";
       inputLote.onkeyup = filtrarTabelaAvancado;
       inputLote.style.width = "70px";
       fAdmin.appendChild(inputLote);
    }

    const inputsFiltro = fAdmin.querySelectorAll("input, select");
    inputsFiltro.forEach(el => {
       el.style.width = el.id === "fNome" ? "200px" : el.id === "fLote" ? "70px" : "auto";
       el.style.padding = "5px";
       el.style.fontSize = "12px";
       el.style.border = "1px solid #334155";
       el.style.borderRadius = "4px";
       el.style.background = "#1e293b";
       el.style.color = "white";
    });
  }
  
  const cabecalho = document.getElementById("cabecalhoTabela");
  const colNames = ["ID", "CPF", "NOME", "NASC", "MUNICIPIO", "TEL", "VIA", "PARCEIRO", "DATA", "ATENDENTE", "BOLETO", "STATUS", "MOTIVO", "DATA STATUS", "NUM CARTEIRA", "LOTE", "AÇÕES"];
  
  cabecalho.innerHTML = colNames.map((name, idx) => `<th class="col-${idx}">${name}</th>`).join("");

  if(!document.getElementById("containerChecks")){
    const divChecks = document.createElement("div");
    divChecks.id = "containerChecks";
    divChecks.style = "display:flex; flex-wrap:wrap; gap:10px; padding:10px; background:#1e293b; border-radius:8px; margin-bottom:10px; font-size:11px; color:#22c55e; border:1px solid #334155;";
    divChecks.innerHTML = "<div style='width:100%; color:white; font-weight:bold; margin-bottom:5px;'>Exibir/Ocultar Colunas:</div>";
    colNames.forEach((name, idx) => {
      divChecks.innerHTML += `<label style="cursor:pointer;"><input type="checkbox" checked onclick="alternarColuna(${idx})"> ${name}</label>`;
    });
    
    divChecks.innerHTML += `<button onclick="fecharLotePorParceiro()" style="margin-left:auto; background:#ef4444; color:white; border:none; padding:5px 10px; border-radius:4px; cursor:pointer; font-weight:bold;">FECHAR LOTE</button>`;
    
    document.getElementById("listasBox").prepend(divChecks);
  }

  const tbody = document.getElementById("corpoTabelaListas");
  tbody.innerHTML = "<tr><td colspan='17'>Carregando dados...</td></tr>";
  
  try {
    const dados = await chamarAPI({ acao: "obterListaCadastros", parceiro: user.parceiro });
    tbody.innerHTML = "";
    
    dados.forEach(item => {
      let valDataStatus = "";
      for (let key in item) {
        let normalizedKey = key.toUpperCase().replace(/\s|_/g, "");
        if (normalizedKey === "DATASTATUS") { valDataStatus = item[key]; break; }
      }
      
      let valTel = "";
      for (let key in item) {
        let normalizedKey = key.toUpperCase().replace(/\s|_/g, "");
        if (normalizedKey === "TEL" || normalizedKey === "TELEFONE") { valTel = item[key]; break; }
      }

      tbody.innerHTML += `<tr>
        <td class="col-0">${item.id || ''}</td>
        <td class="col-1">${item.cpf || ''}</td>
        <td class="col-2">${item.nome || ''}</td>
        <td class="col-3">${item.nasc || ''}</td>
        <td class="col-4">${item.municipio || ''}</td>
        <td class="col-5">${valTel}</td>
        <td class="col-6">${item.via || ''}</td>
        <td class="col-7">${item.parceiro || ''}</td>
        <td class="col-8">${item.data || ''}</td>
        <td class="col-9">${item.atendente || ''}</td>
        <td class="col-10">${item.boleto || ''}</td>
        <td class="col-11"><b>${item.status || ''}</b></td>
        <td class="col-12">${item.motivo || ''}</td>
        <td class="col-13">${valDataStatus}</td>
        <td class="col-14">${item.carteira || ''}</td>
        <td class="col-15">${item.lote || ''}</td>
        <td class="col-16">
          <button onclick='prepararEdicao(${JSON.stringify(item)})' style="background:#f59e0b; color:white; border:none; padding:3px 8px; border-radius:4px; cursor:pointer;">Editar</button>
        </td>
      </tr>`;
    });
    
    const checks = document.getElementById("containerChecks").querySelectorAll("input");
    checks.forEach((chk, i) => { if(!chk.checked) aplicarOcultacao(i, false); });
  } catch (err) {
    tbody.innerHTML = "<tr><td colspan='17'>Erro ao carregar lista.</td></tr>";
  }
}

function alternarColuna(idx) { aplicarOcultacao(idx, event.target.checked); }
function aplicarOcultacao(idx, exibir) {
  document.querySelectorAll(`.col-${idx}`).forEach(c => c.style.display = exibir ? "" : "none");
}

async function fecharLotePorParceiro() {
  const user = JSON.parse(sessionStorage.getItem("usuario"));
  if(!confirm("Deseja fechar o lote atual para o parceiro " + user.parceiro + "?")) return;
  
  try {
    const res = await chamarAPI({ acao: "fecharLoteAppsScript", parceiro: user.parceiro });
    if(res.sucesso) {
      alert("Lote fechado com sucesso! Lote: " + res.loteGerado);
      carregarLista();
    } else {
      alert("Erro ao fechar lote: " + res.erro);
    }
  } catch (err) {
    alert("Erro de conexão ao fechar lote.");
  }
}

async function salvarEntrega() {
  const ctr = document.getElementById("codigoCtr").value;
  const isTerceiro = document.getElementById("checkTerceiro").checked;
  
  if(!alunoEncontradoGlobal) { 
    alert("Informe um CTR válido e aguarde a busca."); 
    return; 
  }
  
  const nomeRecebedor = isTerceiro ? document.getElementById("nomeTerceiro").value : alunoEncontradoGlobal.nome;
  const cpfRecebedor = isTerceiro ? document.getElementById("cpfTerceiro").value : alunoEncontradoGlobal.cpf;
  const vinculo = isTerceiro ? document.getElementById("parentesco").value : "Titular";
  const via = document.querySelector('input[name="viaEntrega"]:checked').value;

  if(isTerceiro && (!nomeRecebedor || !cpfRecebedor || !vinculo)) { 
    alert("Preencha todos os campos do recebedor!"); 
    return; 
  }

  const user = JSON.parse(sessionStorage.getItem("usuario"));

  try {
    const res = await chamarAPI({
      acao: "registrarEntregaAppsScript",
      ctr: ctr,
      cpfA: alunoEncontradoGlobal.cpf,
      nomeA: alunoEncontradoGlobal.nome,
      cpfR: cpfRecebedor,
      nomeR: nomeRecebedor,
      vinculo: vinculo,
      atendente: user.nome,
      parceiro: user.parceiro,
      via: via
    });

    if(res.sucesso) {
      imprimirProtocoloEntrega(ctr, alunoEncontradoGlobal.nome, alunoEncontradoGlobal.cpf, nomeRecebedor, cpfRecebedor, vinculo, user.nome, via);
      alert("Entrega realizada com sucesso!");

      document.getElementById("codigoCtr").value = "";
      document.getElementById("infoAlunoEntrega").style.display = "none";
      if(isTerceiro) {
        document.getElementById("nomeTerceiro").value = "";
        document.getElementById("cpfTerceiro").value = "";
        document.getElementById("parentesco").value = "";
        document.getElementById("checkTerceiro").checked = false;
        if(typeof toggleTerceiro === "function") toggleTerceiro();
      }
      document.getElementById("via1").checked = true;
      alunoEncontradoGlobal = null;
    } else {
      alert("Erro ao salvar: " + res.erro);
    }
  } catch (err) {
    alert("Erro de conexão ao registrar entrega.");
  }
}

document.addEventListener('blur', async function(e){
  if(e.target.id === "codigoCtr"){
    const ctr = e.target.value.trim();
    if(!ctr) return;

    const userStr = sessionStorage.getItem("usuario");
    if(!userStr) return;
    const user = JSON.parse(userStr);

    try {
      const aluno = await chamarAPI({ 
        acao: "buscarPorCodigoAppsScript", 
        ctr: ctr, 
        parceiro: user.parceiro 
      });
      
      if(aluno && aluno.encontrado) {
        alunoEncontradoGlobal = aluno;
        document.getElementById("resNomeAluno").innerText = aluno.nome;
        document.getElementById("resCpfAluno").innerText = aluno.cpf; 
        document.getElementById("infoAlunoEntrega").style.display = "block";
        if(aluno.via) {
          const viaRadio = document.getElementById("via" + aluno.via);
          if(viaRadio) viaRadio.checked = true;
        }
      } else {
        document.getElementById("infoAlunoEntrega").style.display = "none";
        alunoEncontradoGlobal = null;
      }
    } catch (err) {
      console.error("Erro na busca por CTR:", err);
    }
  }
}, true);

function imprimirProtocoloEntrega(ctr, aluno, cpfA, recebedor, cpfR, vinculo, atendente, via) {
  const telaPrint = window.open('', '_blank');
  const dataHora = new Date().toLocaleString('pt-BR');

  telaPrint.document.write(`
    <html>
    <head>
      <title>Entrega CTR ${ctr}</title>
      <style>
        @page { size: A5 landscape; margin: 0; }
        body { font-family: Arial, sans-serif; margin: 0; padding: 7mm; box-sizing: border-box; width: 210mm; height: 148mm; display: flex; justify-content: center; align-items: center; }
        .ticket { width: 100%; height: 100%; border: 2px solid #000; padding: 6mm; box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between; }
        .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 2mm; }
        .header h2 { margin: 0; font-size: 20px; text-transform: uppercase; }
        .section-title { font-size: 13px; font-weight: bold; background: #f2f2f2; padding: 1.5mm 3mm; border: 1px solid #000; margin-top: 2mm; }
        .info-group { display: flex; width: 100%; font-size: 14px; margin: 1.5mm 0; }
        .info-item { flex: 1; line-height: 1.4; }
        .declaracao { font-size: 13px; line-height: 1.5; text-align: justify; border: 1px dashed #000; padding: 4mm; margin: 3mm 0; }
        .assinatura-area { text-align: center; margin-top: 2mm; }
        .assinatura-linha { border-top: 1.5px solid #000; width: 60%; margin: 0 auto; }
        .assinatura-texto { font-size: 11px; margin-top: 1mm; font-weight: bold; text-transform: uppercase; }
        .meta-info { display: flex; justify-content: space-between; font-size: 10px; border-top: 1px solid #eee; padding-top: 1mm; }
        b { text-transform: uppercase; }
      </style>
    </head>
    <body>
      <div class="ticket">
        <div class="header"><h2>COMPROVANTE DE ENTREGA</h2></div>
        <div>
          <div class="section-title">DADOS DO ALUNO</div>
          <div class="info-group">
            <div class="info-item"><b>CTR:</b> ${ctr}</div>
            <div class="info-item"><b>VIA:</b> ${via}ª VIA</div>
            <div class="info-item"><b>ALUNO:</b> ${aluno}</div>
          </div>
          <div class="info-group" style="margin-top: -1mm;"><div class="info-item"><b>CPF ALUNO:</b> ${cpfA}</div></div>
        </div>
        <div class="declaracao">Declaro que recebi, nesta data, a Carteira de Estudante Macrorregião 2026...</div>
        <div>
          <div class="section-title">DADOS DO RECEBEDOR</div>
          <div class="info-group">
            <div class="info-item"><b>NOME:</b> ${recebedor}</div>
            <div class="info-item"><b>CPF:</b> ${cpfR}</div>
          </div>
          <div class="info-group" style="margin-top: -1mm;"><div class="info-item"><b>VÍNCULO:</b> ${vinculo}</div></div>
        </div>
        <div class="assinatura-area"><div class="assinatura-linha"></div><div class="assinatura-texto">Assinatura do Recebedor</div></div>
        <div class="meta-info"><span><b>ATENDENTE:</b> ${atendente}</span><span><b>DATA/HORA:</b> ${dataHora}</span></div>
      </div>
      <script>window.onload = function() { window.print(); window.onafterprint = function() { window.close(); }; };<\/script>
    </body>
    </html>
  `);
  telaPrint.document.close();
}

function abrirAdmin() {
  var urlLotes = "https://script.google.com/macros/s/AKfycbxeyoKG99zETrrx6BdF7--w_-1cVe-S0tctxKOAfgFFQ3_as64oRqONoditWtXWsrRF/exec";
  var token = "MACRO@MACRO";
  var u = ""; var s = "";
  try {
    var uField = document.getElementById('userLogin');
    var sField = document.getElementById('passLogin');
    if (uField) u = uField.value;
    if (sField) s = sField.value;
  } catch (e) {}
  var linkFinal = urlLotes + "?u=" + encodeURIComponent(u) + "&s=" + encodeURIComponent(s) + "&token=" + encodeURIComponent(token);
  window.open(linkFinal, '_blank');
}

function mascaraData(campo) {
  var v = campo.value.replace(/\D/g, "");
  if (v.length >= 2) v = v.substring(0, 2) + "/" + v.substring(2);
  if (v.length >= 5) v = v.substring(0, 5) + "/" + v.substring(5, 9);
  campo.value = v;
}
