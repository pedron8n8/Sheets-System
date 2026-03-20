# ============================================================================
# STREAMLIT APP - Property Analysis Form
# ============================================================================
# Este app permite que usuários preencham um formulário de análise de 
# propriedades imobiliárias e salve os dados em um arquivo Excel.

import streamlit as st          # Framework para criar interface web
import pandas as pd             # Biblioteca para trabalhar com dados (Excel, CSV, etc)
import os                       # Biblioteca para verificar se arquivos existem
from datetime import datetime   # Para trabalhar com datas
from main import processar_registro_por_indice

# ============================================================================
# CONFIGURAÇÕES INICIAIS
# ============================================================================

# Define o nome do arquivo Excel onde os dados serão salvos
ARQUIVO_EXCEL = 'contatos.xlsx'

# Define o título da página que aparece no início da interface
st.title("🚀 Property Analysis Form")

# ============================================================================
# INICIALIZAÇÃO DO ESTADO DA SESSÃO
# ============================================================================
# O st.session_state mantém dados durante toda a sessão do usuário.
# Sem ele, os dados desapareceriam quando a página é atualizada.

# Inicializa a lista de receitas adicionais (Other Income) se não existir
# Esta lista será usada para armazenar as linhas dinâmicas que o usuário adiciona
if 'campos_extras' not in st.session_state:
    st.session_state.campos_extras = []

# Inicializa a lista de despesas adicionais (Other Expenses) se não existir
# Esta lista será usada para armazenar as linhas dinâmicas que o usuário adiciona
if 'despesas_extras' not in st.session_state:
    st.session_state.despesas_extras = []

# ============================================================================
# FUNÇÕES PARA ADICIONAR CAMPOS DINÂMICOS
# ============================================================================

# Função que adiciona uma nova linha de receita (Other Income)
def adicionar_campo():
    # Adiciona um novo dicionário com label (nome) e valor (amount) vazios
    st.session_state.campos_extras.append({"label_selecionada": "", "valor": ""})

# Função que adiciona uma nova linha de despesa (Other Expense)
def adicionar_despesa():
    # Adiciona um novo dicionário com label (nome) e valor (amount) vazios
    st.session_state.despesas_extras.append({"label": "", "valor": ""})

def remover_campo(indice):
    # Remove uma linha de receita e reorganiza as chaves dinâmicas do session_state
    if 0 <= indice < len(st.session_state.campos_extras):
        tamanho_antigo = len(st.session_state.campos_extras)
        st.session_state.campos_extras.pop(indice)

        for j in range(indice, tamanho_antigo - 1):
            for prefixo in ["select_", "input_nome_", "input_valor_"]:
                chave_origem = f"{prefixo}{j+1}"
                chave_destino = f"{prefixo}{j}"
                if chave_origem in st.session_state:
                    st.session_state[chave_destino] = st.session_state[chave_origem]
                else:
                    st.session_state.pop(chave_destino, None)

        for prefixo in ["select_", "input_nome_", "input_valor_"]:
            st.session_state.pop(f"{prefixo}{tamanho_antigo - 1}", None)

def remover_despesa(indice):
    # Remove uma linha de despesa e reorganiza as chaves dinâmicas do session_state
    if 0 <= indice < len(st.session_state.despesas_extras):
        tamanho_antigo = len(st.session_state.despesas_extras)
        st.session_state.despesas_extras.pop(indice)

        for j in range(indice, tamanho_antigo - 1):
            for prefixo in ["sel_exp_", "txt_exp_nome_", "txt_exp_val_"]:
                chave_origem = f"{prefixo}{j+1}"
                chave_destino = f"{prefixo}{j}"
                if chave_origem in st.session_state:
                    st.session_state[chave_destino] = st.session_state[chave_origem]
                else:
                    st.session_state.pop(chave_destino, None)

        for prefixo in ["sel_exp_", "txt_exp_nome_", "txt_exp_val_"]:
            st.session_state.pop(f"{prefixo}{tamanho_antigo - 1}", None)

# ============================================================================
# CRIANDO O FORMULÁRIO PRINCIPAL
# ============================================================================
# O st.form() cria um formulário que é submetido completamente ao clicar
# o botão de submit. Sem o form, cada widget atualizaria a página individualmente.
# clear_on_submit=True limpa os campos após o envio.

delete_income_index = None
delete_expense_index = None

with st.form("meu_formulario", clear_on_submit=False, enter_to_submit=False):
    
    # ========== SEÇÃO 1: INFORMAÇÕES DA PROPRIEDADE ==========
    st.markdown("### PROPERTY INFORMATION")
    
    # Campo de texto simples: Nome da propriedade
    propertyName = st.text_input("Property Name")
    
    # Campo de texto simples: Tipo de propriedade (apartment, house, commercial, etc)
    propertyType = st.text_input("Property Type")
    
    # Campo de texto simples: Endereço da propriedade
    adreess = st.text_input("Address")
    
    # Campo de texto simples: Cidade e estado
    cityAndState = st.text_input("City and State")
    
    # Campo numérico: Quantidade de unidades (mín 0, incrementa de 1 em 1)
    numberOfUnits = st.number_input("Number of Units", min_value=0, step=1)
    
    # Campo numérico com moeda: Preço de compra (formato com 2 casas decimais)
    purchasePrice = st.number_input("Purchase Price", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico em percentual: Entrada (Down Payment)
    downPayment = st.number_input("Down Payment (%)", min_value=0.0, max_value=100.0, step=0.1, format="%.2f")
    
    # Campo numérico em percentual: Custos de due diligence 
    dueDiligenceCosts = st.number_input(
        "Due Diligence Costs", 
        min_value=0.0, 
        step=100.0,
        format="%.2f"     # Exibe sempre 2 casas decimais
    )
    
    # Campo numérico com moeda: Custos de origination do empréstimo
    loanOriginalCosts = st.number_input("Loan Original Costs %", min_value=0.0, max_value=100.0, step=0.1, format="%.2f")
    
    # Campo de data: Data da compra (começa com data de hoje)
    purchaseDate = st.date_input("Purchase Date", value=datetime.today())

    # Campo numérico inteiro: horizonte do modelo em anos
    endYear = st.number_input("End Year", min_value=1, step=1, value=10)

    # ========== SEÇÃO 2: SUPOSIÇÕES DE RECEITA ==========
    st.markdown("### REVENUE ASSUMPTIONS")
    
    # Campo numérico com moeda: Aluguel potencial bruto
    grossPotentialRent = st.number_input("Gross Potential Rent", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico em percentual: Taxa de vacância (0-100%, começa em 5%)
    vacancyRate = st.number_input(
        "Vacancy Rate %",
        min_value=0.0,
        max_value=100.0,
        value=5.0,          # Valor padrão: 5% (comum no mercado)
        step=0.1,
        format="%.2f"
    )
    
    # Campo numérico em percentual: Perda por crédito (0-100%, começa em 1%)
    creditLoss = st.number_input(
        "Credit Loss %",
        min_value=0.0,
        max_value=100.0,
        value=1.0,
        step=0.1,
        format="%.2f"
    )

    # ========== SEÇÃO 2B: RECEITAS ADICIONAIS (DINÂMICAS) ==========
    # Esta seção exibe linhas que o usuário pode adicionar clicando no botão
    
    # Verifica se há receitas adicionais na lista
    if st.session_state.campos_extras:
        st.markdown("#### Other Income")
        
        # Itera sobre cada receita adicional que foi adicionada
        for i, campo in enumerate(st.session_state.campos_extras):
            # Cria 3 colunas: tipo, valor e botão de excluir
            col1, col2, col3 = st.columns([2, 1, 0.7])
            
            # -------- COLUNA 1: Seleção de tipo de receita --------
            with col1:
                # Lista de opções pré-definidas para o usuário escolher
                opcoes = [
                    "Select...", 
                    "Other Income - Parking",
                    "Other Income - Laundry",
                    "Other Income - Misc",
                    "Other (Type name...)"
                ]
                
                # Evita opções duplicadas entre linhas (exceto "Selecione" e "Outro")
                escolha_atual = st.session_state.get(f"select_{i}", "Select...")
                escolhidas_outras_linhas = {
                    st.session_state.get(f"select_{j}")
                    for j in range(len(st.session_state.campos_extras))
                    if j != i
                }
                opcoes_disponiveis = [
                    opcao
                    for opcao in opcoes
                    if (
                        opcao in ["Select...", "Other (Type name...)"]
                        or opcao == escolha_atual
                        or opcao not in escolhidas_outras_linhas
                    )
                ]

                indice_padrao = (
                    opcoes_disponiveis.index(escolha_atual)
                    if escolha_atual in opcoes_disponiveis
                    else 0
                )

                # Cria um selectbox para escolher a receita
                escolha = st.selectbox(
                    f"Income Type {i+1}",
                    opcoes_disponiveis,
                    index=indice_padrao,
                    key=f"select_{i}"
                )

                # Mantem o campo customizado sempre visivel para evitar travas de edicao em formularios.
                nome_personalizado = st.text_input(
                    f"Income Name {i+1} (optional)",
                    key=f"input_nome_{i}",
                )

                if escolha == "Other (Type name...)":
                    st.session_state.campos_extras[i]['label_selecionada'] = nome_personalizado.strip()
                else:
                    # Caso contrario, usa o valor selecionado
                    st.session_state.campos_extras[i]['label_selecionada'] = escolha

            # -------- COLUNA 2: Campo para o valor da receita --------
            with col2:
                # Campo de texto para inserir o valor (em formato livre)
                valor = st.text_input(f"Amount USD {i+1}", key=f"input_valor_{i}")
                # Armazena o valor no estado da sessão
                st.session_state.campos_extras[i]['valor'] = valor

            with col3:
                # Espaço para alinhar o botão com os campos (label + input)
                st.markdown("<div style='height: 1.9rem;'></div>", unsafe_allow_html=True)
                if st.form_submit_button(
                    "🗑️ Delete",
                    key=f"del_income_{i}",
                    use_container_width=True,
                    type="secondary"
                ):
                    delete_income_index = i
        
        # Dentro de st.form, use sempre st.form_submit_button (não st.button)
        add_income_clicked = st.form_submit_button("➕ Add Other Income", key="btn_add_income")
    else:
        # Se não há receitas adicionais ainda, mostra apenas o botão
        add_income_clicked = st.form_submit_button("➕ Add Other Income", key="btn_add_income")

    # ========== SEÇÃO 3: SUPOSIÇÕES DE DESPESAS ==========
    st.markdown("### EXPENSE ASSUMPTIONS")
    
    # Campo numérico com moeda: Imposto sobre propriedade
    propertyTax = st.number_input("Property Tax", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico com moeda: Seguro da propriedade
    insurance = st.number_input("Insurance", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico em percentual: Taxa de gerenciamento da propriedade
    # (%EGI = porcentagem da receita bruta efetiva)
    managementFee = st.number_input("Property Management Fee (%EGI)", min_value=0.0, max_value=100.0, value=5.0, step=0.1, format="%.2f")
    
    # Campo numérico com moeda: Reparos e manutenção
    repairsAndMaintenance = st.number_input("Repairs and Maintenance", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico com moeda: Serviços de utilidade (água, gás, eletricidade)
    utilities = st.number_input("Utilities", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico com moeda: Despesas de capital (investimentos em melhorias)
    capitalExpenditures = st.number_input("Capital Expenditures", min_value=0.0, step=100.0, format="%.2f")
    
    # Campo numérico com moeda: Limpeza e paisagismo
    landscapeAndJanitorial = st.number_input("Landscape and Janitorial", min_value=0.0, step=100.0, format="%.2f")

    # ========== SEÇÃO 3B: DESPESAS ADICIONAIS (DINÂMICAS) ==========
    # Esta seção funciona de forma similar às receitas adicionais
    
    if st.session_state.despesas_extras:
        st.markdown("#### Other Expenses")
        
        # Itera sobre cada despesa adicional que foi adicionada
        for i, despesa in enumerate(st.session_state.despesas_extras):
            # Cria 3 colunas: tipo, valor e botão de excluir
            col1, col2, col3 = st.columns([2, 1, 0.7])
            
            # -------- COLUNA 1: Seleção de tipo de despesa --------
            with col1:
                # Lista de opções pré-definidas para despesas
                opcoes_exp = [
                    "Select...",
                    "Marketing & Advertising",      # Publicidade e marketing
                    "Reserves for Replacement",     # Reserva para reposição
                    "Trash Removal",                # Coleta de lixo
                    "Pest Control",                 # Controle de pragas
                    "Security / Access Control",    # Segurança e controle de acesso
                    "Other Expense (Type...)"       # Opção customizada
                ]
                
                # Evita opções duplicadas entre linhas (exceto "Selecione" e "Other Expense")
                escolha_exp_atual = st.session_state.get(f"sel_exp_{i}", "Select...")
                escolhidas_exp_outras_linhas = {
                    st.session_state.get(f"sel_exp_{j}")
                    for j in range(len(st.session_state.despesas_extras))
                    if j != i
                }
                opcoes_exp_disponiveis = [
                    opcao
                    for opcao in opcoes_exp
                    if (
                        opcao in ["Select...", "Other Expense (Type...)"]
                        or opcao == escolha_exp_atual
                        or opcao not in escolhidas_exp_outras_linhas
                    )
                ]

                indice_exp_padrao = (
                    opcoes_exp_disponiveis.index(escolha_exp_atual)
                    if escolha_exp_atual in opcoes_exp_disponiveis
                    else 0
                )

                # Selectbox para escolher a despesa
                escolha = st.selectbox(
                    f"Expense Type {i+1}",
                    opcoes_exp_disponiveis,
                    index=indice_exp_padrao,
                    key=f"sel_exp_{i}"
                )

                # Mantem o campo customizado sempre visivel para evitar travas de edicao em formularios.
                nome_custom = st.text_input(
                    f"Expense Name {i+1} (optional)",
                    key=f"txt_exp_nome_{i}",
                )

                if escolha == "Other Expense (Type...)":
                    st.session_state.despesas_extras[i]['label'] = nome_custom.strip()
                else:
                    st.session_state.despesas_extras[i]['label'] = escolha

            # -------- COLUNA 2: Campo para o valor da despesa --------
            with col2:
                # Campo de texto para inserir o valor da despesa
                valor_exp = st.text_input(f"Amount USD {i+1}", key=f"txt_exp_val_{i}")
                st.session_state.despesas_extras[i]['valor'] = valor_exp

            with col3:
                # Espaço para alinhar o botão com os campos (label + input)
                st.markdown("<div style='height: 1.9rem;'></div>", unsafe_allow_html=True)
                if st.form_submit_button(
                    "🗑️ Delete",
                    key=f"del_expense_{i}",
                    use_container_width=True,
                    type="secondary"
                ):
                    delete_expense_index = i
        
        # Dentro de st.form, use sempre st.form_submit_button (não st.button)
        add_expense_clicked = st.form_submit_button("➕ Add Other Expense", key="btn_add_expense")
    else:
        # Se não há despesas adicionais ainda, mostra apenas o botão
        add_expense_clicked = st.form_submit_button("➕ Add Other Expense", key="btn_add_expense")

    # ========== SEÇÃO 4: DESPESAS DE CAPITAL (CAPEX) ==========
    st.markdown("### CAPITAL EXPENDITURES")
    
    # Campos de texto para registrar até 5 itens de despesas de capital
    # Estes são inputs livres, sem sugestões pré-definidas
    caPex1 = st.text_input("CapEx 1", key="capex1")
    caPex2 = st.text_input("CapEx 2", key="capex2")
    caPex3 = st.text_input("CapEx 3", key="capex3")
    caPex4 = st.text_input("CapEx 4", key="capex4")
    caPex5 = st.text_input("CapEx 5", key="capex5")

    # ========== BOTÃO DE ENVIO DO FORMULÁRIO ==========
    # Este botão submete todo o formulário de uma vez
    botao_salvar = st.form_submit_button("💾 Save")

# Ações dos botões auxiliares do formulário
if delete_income_index is not None:
    remover_campo(delete_income_index)
    st.rerun()

if delete_expense_index is not None:
    remover_despesa(delete_expense_index)
    st.rerun()

if add_income_clicked:
    adicionar_campo()
    st.rerun()

if add_expense_clicked:
    adicionar_despesa()
    st.rerun()

# ============================================================================
# PROCESSAR O ENVIO DO FORMULÁRIO
# ============================================================================
# Quando o usuário clica em "Salvar na Planilha", botao_salvar recebe True
# Então executamos a lógica de salvamento dos dados

if botao_salvar:
    # Verifica se os campos obrigatórios foram preenchidos
    # (Property Name e Property Type são necessários)
    if propertyName and propertyType:
        # Monta registro base com os campos fixos do formulário
        registro = {
            'Property Name': propertyName,
            'Property Type': propertyType,
            'Address': adreess,
            'City and State': cityAndState,
            'Number of Units': numberOfUnits,
            'Purchase Price': purchasePrice,
            'Down Payment (%)': downPayment,
            'Due Diligence Costs': dueDiligenceCosts,
            'Loan Original Costs': loanOriginalCosts,
            'Purchase Date': purchaseDate,
            'End Year': int(endYear),
            'Gross Potential Rent': grossPotentialRent,
            'Vacancy Rate %': vacancyRate,
            'Credit Loss %': creditLoss,
            'Property Tax': propertyTax,
            'Insurance': insurance,
            'Management Fee %': managementFee,
            'Repairs and Maintenance': repairsAndMaintenance,
            'Utilities': utilities,
            'Capital Expenditures': capitalExpenditures,
            'Landscape and Janitorial': landscapeAndJanitorial,
            'CapEx 1': caPex1,
            'CapEx 2': caPex2,
            'CapEx 3': caPex3,
            'CapEx 4': caPex4,
            'CapEx 5': caPex5,
            'Submitted': 'No'
        }

        # Salva receitas adicionais em colunas separadas na tabela
        income_count = 0
        for item in st.session_state.campos_extras:
            label = str(item.get('label_selecionada', '')).strip()
            valor = str(item.get('valor', '')).strip()
            if label and label != 'Select...' and valor:
                income_count += 1
                registro[f'Other Income {income_count} Type'] = label
                registro[f'Other Income {income_count} Amount'] = valor

        # Salva despesas adicionais em colunas separadas na tabela
        expense_count = 0
        for item in st.session_state.despesas_extras:
            label = str(item.get('label', '')).strip()
            valor = str(item.get('valor', '')).strip()
            if label and label != 'Select...' and valor:
                expense_count += 1
                registro[f'Other Expense {expense_count} Type'] = label
                registro[f'Other Expense {expense_count} Amount'] = valor

        registro['Other Income Count'] = income_count
        registro['Other Expense Count'] = expense_count
        
        # -------- PASSO 1: CRIAR DATAFRAME COM OS DADOS --------
        # pd.DataFrame cria uma tabela com os dados coletados do formulário
        # Cada coluna representa um campo que o usuário preencheu
        novos_dados = pd.DataFrame([registro])

        # -------- PASSO 2: ANEXAR AOS DADOS EXISTENTES OU CRIAR NOVO --------
        # Verifica se o arquivo Excel já existe
        if os.path.exists(ARQUIVO_EXCEL):
            # Se existe: ler dados antigos, concatenar com novos dados
            df_antigo = pd.read_excel(ARQUIVO_EXCEL)           # Lê o arquivo existente
            # Normaliza colunas antigas para o schema atual
            legacy_map = {
                'Down Payment': 'Down Payment (%)',
                'Due Diligence Costs %': 'Due Diligence Costs',
            }
            for old_col, new_col in legacy_map.items():
                if old_col in df_antigo.columns:
                    if new_col in df_antigo.columns:
                        df_antigo[new_col] = df_antigo[new_col].combine_first(df_antigo[old_col])
                        df_antigo = df_antigo.drop(columns=[old_col])
                    else:
                        df_antigo = df_antigo.rename(columns={old_col: new_col})
            # Garante que coluna "Submitted" existe (compatibilidade com arquivos antigos)
            if 'Submitted' not in df_antigo.columns:
                df_antigo['Submitted'] = 'No'
            df_final = pd.concat([df_antigo, novos_dados], ignore_index=True)  # Junta os dois
        else:
            # Se não existe: os novos dados serão o início do arquivo
            df_final = novos_dados

        # -------- PASSO 3: SALVAR DE VOLTA PARA O ARQUIVO EXCEL --------
        # Salva o DataFrame (com dados antigos + novos) de volta ao Excel
        df_final.to_excel(ARQUIVO_EXCEL, index=False)

        # Processa imediatamente o registro novo e gera a planilha de output.
        novo_indice = len(df_final) - 1
        caminho_gerado = processar_registro_por_indice(novo_indice)
        
        # Exibe mensagem de sucesso para o usuário
        st.success(f"✅ Property '{propertyName}' saved successfully!")
        if caminho_gerado:
            st.info(f"📁 Output gerado: {caminho_gerado}")
        else:
            st.warning("⚠️ Registro salvo, mas nao foi possivel gerar o Output automaticamente.")
    else:
        # Se os campos obrigatórios não foram preenchidos, mostra erro
        st.error("⚠️ Please fill in at least Property Name and Property Type.")

# ============================================================================
# EXIBIR OS DADOS SALVOS
# ============================================================================
# Mostra a tabela com todos os dados que foram salvos até agora

# if os.path.exists(ARQUIVO_EXCEL):
#     st.markdown("---")                  # Cria uma linha divisória visual
#     st.write("### 📊 Dados Atuais")    # Título da seção
#     # Lê o arquivo Excel e exibe como tabela interativa
#     # use_container_width=True faz a tabela ocupar toda a largura da página
#     st.dataframe(pd.read_excel(ARQUIVO_EXCEL), use_container_width=True)