/*Globais Auxiliares*/
const codForm = ProcessData.processId
const codVersao = ProcessData.version
const codProcesso = ProcessData.processInstanceId
const codEtapa = ProcessData.activityInstanceId
const codCiclo = ProcessData.cycle
const tituloProcesso = ProcessData.title.replace(/\n /g, '')
const tituloEtapa = ProcessData.activityTitle.replace(/\n /g, '')
const BPM_URL = location.protocol + '//' + location.hostname + '/bpm'

/*Etapas*/
var Etapa = Object.freeze({
    SOLICITAR_ATENDIMENTO_FISCAL: 3,
    REALIZAR_ATENDIMENTO: 4,
    ENCERRAR_PROCESSO: 2,
    CANCELAR: 1
})

$(document).ready(function () {
    initForm();
    setForm();
});

/* inicialização das atividades */
function initForm() {
    //Switch case com as etapas declaradas na variavel etapa (Object.freeze)
    //Ex:

    switch (codEtapa) {

        case Etapa.SOLICITAR_ATENDIMENTO_FISCAL:
            bloquearAdicionarLinhaSeNecessario();
            controlaVisibilidadeGrupo3();
            break
        case Etapa.REALIZAR_ATENDIMENTO:
            bloquearAdicionarLinhaSeNecessario();
            break
        default:
            break
    }
}

/* ações dos campos por atividade */
function setForm() {
    switch (codEtapa) {
        case Etapa.SOLICITAR_ATENDIMENTO_FISCAL:

            Form.grids('DADOS_NF').fields("LT_CNPJ_CPF_EMITENTE").subscribe("BLUR", function () {
                console.log("Blur disparado");
                consultarFornecedorFrisia();
            });

            Form.fields('LS_CLASSIFICACAO').subscribe('SET_FIELD_VALUE', function () {
                atualizarObrigatoriedadeJustificativa();
            });

            Form.grids('DADOS_NF').fields('LS_STATUS_OC').subscribe('SET_FIELD_VALUE', function () {
                RegraCamposGrid();
            });

            Form.actions("aprovar").subscribe("SUBMIT", function (itemId, action, reject) {
                var valido = validarQuantidadeLinhasGrid();

                if (!valido) {
                    reject();
                    return;
                }

                var errors = Form.errors();
                if (Object.keys(errors).length)
                    reject();
            });

            Form.grids('DADOS_NF').subscribe('GRID_ADD_BEFORE', function (formId, gridId, submittedDataRow, dataRows, resolve, reject) {
                if (dataRows.length >= 10) {
                    reject("A tabela permite no máximo 10 linhas. Remova uma linha para adicionar outra.");
                } else {
                    resolve();
                }
            });

            Form.grids('DADOS_NF').subscribe('GRID_ADD_AFTER', function () {
                bloquearAdicionarLinhaSeNecessario();
            });

            Form.grids('DADOS_NF').subscribe('GRID_DELETE_AFTER', function () {
                bloquearAdicionarLinhaSeNecessario();
            });

            break;
        case Etapa.REALIZAR_ATENDIMENTO:

            Form.fields('RB_DADOS_NF').subscribe('CHANGE', function () {
                controlaJustificativaFiscal();
            });

            Form.actions("rejeitar").subscribe("SUBMIT", function (formId, actionId, reject) {
                const campoRbDados = Form.fields('RB_DADOS_NF');
                const valorRbDados = campoRbDados.value();

                if (valorRbDados !== "Não") {
                    campoRbDados.errors("Para rejeitar, o campo deve estar marcado como 'Não'.").apply();
                    reject(); // impede o submit
                } else {
                    // Limpa erros, garante que está tudo certo
                    campoRbDados.errors([]).apply();
                    // NÃO chama reject() — o form será submetido normalmente
                }
            });

            break;
        default:
            break
    }
}

function atualizarObrigatoriedadeJustificativa() {
    var classificacaoRaw = Form.fields('LS_CLASSIFICACAO').value();
    var classificacao = Array.isArray(classificacaoRaw) ? classificacaoRaw[0] : classificacaoRaw || "";

    var campoJustificativa = Form.fields('CX_JUST_PRIORIDADE');

    if (classificacao === 'Alta') {
        campoJustificativa.required(true);
    } else {
        campoJustificativa.required(false);
    }

    Form.apply();
}

function RegraCamposGrid() {
    var statusRaw = Form.grids('DADOS_NF').fields('LS_STATUS_OC').value();
    var status = Array.isArray(statusRaw) ? statusRaw[0] : statusRaw || "";

    var num_oc = Form.grids('DADOS_NF').fields('LT_NUMERO_OC');
    var organizacao = Form.grids('DADOS_NF').fields('LS_ORGANIZACAO');
    var saida_oc = Form.grids('DADOS_NF').fields('DC_SALDO_OC');
    var dt_venc = Form.grids('DADOS_NF').fields('DT_VENCIMENTO');

    if (status === 'Aberta') {
        num_oc.disabled(false).required(true);
        organizacao.disabled(false).required(true);
        saida_oc.disabled(false).required(true);
        dt_venc.disabled(false).required(true);
    } else {
        num_oc.disabled(true).required(false);
        organizacao.disabled(true).required(false);
        saida_oc.disabled(true).required(false);
        dt_venc.disabled(true).required(false);
    }

    Form.apply();
}

function validarQuantidadeLinhasGrid() {
    var grid1 = Form.grids('DADOS_NF').dataRows();
    var grid2 = Form.grids('GD_ANEXO').dataRows();

    var erros = {};

    if (grid1.length < 1) {
        erros.DADOS_NF = ['Adicione ao menos um item na grid para seguir'];
    }

    if (grid2.length < 1) {
        erros.GD_ANEXO = ['Adicione ao menos um item na grid para seguir'];
    }

    if (Object.keys(erros).length > 0) {
        Form.errors(erros).apply();
        return false;
    }

    return true;
}

function bloquearAdicionarLinhaSeNecessario() {
    Form.grids("DADOS_NF").columns("LT_TXT_ATENCAO").visible(false).apply();
    var grid = Form.grids('DADOS_NF');
    var qtd = grid.dataRows().length;

    // Seleciona o botão apenas dentro da div com ID 'input__DADOS_NF'
    var $botaoAdicionar = $('#input__DADOS_NF .submit-button-grid');

    if (qtd >= 10) {
        $botaoAdicionar
            .addClass('disabled')
            .css({ 'pointer-events': 'none', 'opacity': 0.5 });
    } else {
        $botaoAdicionar
            .removeClass('disabled')
            .css({ 'pointer-events': 'auto', 'opacity': 1 });
    }
}

function controlaVisibilidadeGrupo3() {

    if (codCiclo > 1) {
        Form.groups('GROUP3').visible(true);
    } else {
        Form.groups('GROUP3').visible(false);
    }
}

function controlaJustificativaFiscal() {
    const rb_dados = Form.fields('RB_DADOS_NF').value();
    const justi = Form.fields('CX_JUST_ATEND_FISCAL');

    if (rb_dados === 'Sim') {
        justi.disabled(true);
        justi.value(''); // limpa o campo
        justi.setRequired('rejeitar', false);
        justi.errors([]).apply(); // limpa erros antigos, se houver
    } else {
        justi.disabled(false).required(true);
        justi.setRequired('rejeitar', true);
        justi.errors([]).apply();
    }

    Form.apply();
}

function consultarFornecedorFrisia() {
    const cnpjField = Form.grids('DADOS_NF').fields("LT_CNPJ_CPF_EMITENTE");
    const razaoField = Form.grids('DADOS_NF').fields("LT_RAZAO_SOCIAL_PRESTAD");
    const tipoField = Form.grids('DADOS_NF').fields("LT_TP");

    const cnpj = cnpjField.value().replace(/\D/g, '');

    if (!cnpj || cnpj.length < 11) {
        cnpjField.errors("CNPJ/CPF inválido").apply();
        razaoField.value('');
        tipoField.value('');
        cnpjField.errors([]).apply();
        Form.apply();
        return;
    } else {
        cnpjField.errors([]).apply();


        Form.waitingForAction(true).apply();

        try {

            var settings = {
                url: "https://extapi.frisia.coop.br/common/SoftExpert/Fornecedor",
                method: "GET",
                headers: {
                    'Authorization': 'Basic c2U6ZnJpc2lhYXBp'
                },
                data: {
                    ds_cpf_cnpj: cnpj
                }
            }

            $.ajax(settings).done(function (response) {
                const fornecedor = response[0];
                razaoField.value(fornecedor.ds_nome).apply();
                tipoField.value(fornecedor.tp_pessoa).apply();
            }).fail(function () {
                console.log("Erro na requisição");
            })

        } catch (error) {
            console.error("Erro ao consultar fornecedor:", error);
            cnpjField.errors("Erro ao consultar fornecedor").apply();
        } finally {
            Form.waitingForAction(false).apply();
        }
    }
}
//if (data.length > 0) {
//
//                } else {
//                    razaoField.value('');
//                    tipoField.value('');
//                    cnpjField.errors("Fornecedor não encontrado").apply();
//                }