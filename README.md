# =====================================
# CONFIGURACOES DA AUTOMACAO
# =====================================

import hashlib
import json
import re
import time
from pathlib import Path
from urllib.parse import urlparse

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

URL_EXCEL = "https://1drv.ms/x/c/fcba0b3173c6e409/IQATzm7u7FECTYwX4fW9vd9ZAazFUAI9U9QirHlz_YrAVYo?e=4Qd4Cy"
NOME_ABA = "Tarefas"

TEMPO_VISUAL = 1.5
DELAY_TAREFA = 0.5
DELAY_FINAL = 1.5

SALVAR_AUTOMATICAMENTE = True
MANTER_NAVEGADOR_ABERTO = True
TEMPO_PARA_LOGIN = 300

PASTA_PERFIL_CHROME = Path(__file__).resolve().with_name("chrome_perfil_excel")

CORES_URGENCIA = {
    5: (234, 153, 153),
    4: (244, 204, 204),
    3: (255, 229, 153),
    2: (182, 215, 168),
    1: (255, 255, 255),
}

STATUS_TAREFA = {
    1: "Não iniciada",
    2: "Em andamento",
    3: "Aguardando terceiros",
    4: "Concluída",
    5: "Suspensa",
}

TAREFAS_PARA_CADASTRAR = [
    {
        "id": "entregar-relatorio",
        "tarefa": "Entregar relatório",
        "tempo_medio": "2h",
        "descricao": "Documentação",
        "realizacoes": 0,
        "status": "Em andamento",
        "ajuda_externa": "Não",
        "urgencia": 5,
    },
    {
        "id": "revisar-projeto",
        "tarefa": "Revisar projeto",
        "tempo_medio": "1h",
        "descricao": "Revisão técnica",
        "realizacoes": 1,
        "status": "Não iniciada",
        "ajuda_externa": "Sim",
        "urgencia": 3,
    },
    {
        "id": "organizar-arquivos",
        "tarefa": "Organizar arquivos",
        "tempo_medio": "30min",
        "descricao": "Organização",
        "realizacoes": 3,
        "status": "Concluída",
        "ajuda_externa": "Não",
        "urgencia": 2,
    },
]

CAMPOS_TAREFA = (
    "tarefa",
    "tempo_medio",
    "descricao",
    "realizacoes",
    "status",
    "ajuda_externa",
    "urgencia",
)


def mostrar_mensagem(texto):
    print(texto)


def validar_configuracoes():
    if not URL_EXCEL or URL_EXCEL == "COLE_AQUI_O_LINK_DO_EXCEL_ONLINE":
        raise ValueError(
            "Configure URL_EXCEL com o link de compartilhamento da pasta de trabalho "
            "do Excel Online."
        )

    partes_url = urlparse(URL_EXCEL)
    host = (partes_url.hostname or "").lower()
    dominios_excel = (
        "1drv.ms",
        "live.com",
        "office.com",
        "officeapps.live.com",
        "onedrive.com",
        "sharepoint.com",
        "microsoft365.com",
    )
    dominio_valido = any(
        host == dominio or host.endswith(f".{dominio}")
        for dominio in dominios_excel
    )
    if partes_url.scheme != "https" or not dominio_valido:
        raise ValueError(
            "URL_EXCEL precisa ser um link HTTPS do Excel Online, OneDrive ou SharePoint."
        )

    if not NOME_ABA.strip():
        raise ValueError("NOME_ABA não pode ficar vazio.")

    for nome, valor in {
        "TEMPO_VISUAL": TEMPO_VISUAL,
        "DELAY_TAREFA": DELAY_TAREFA,
        "DELAY_FINAL": DELAY_FINAL,
    }.items():
        if valor < 0:
            raise ValueError(f"{nome} deve ser maior ou igual a zero.")

    mostrar_mensagem("[1/6] Iniciando automação...")
    mostrar_mensagem("[2/6] Verificando Google Chrome...")


def abrir_excel_online():
    """Abre a pasta de trabalho em uma janela Chrome com perfil persistente."""
    opcoes = webdriver.ChromeOptions()
    opcoes.add_argument(f"--user-data-dir={PASTA_PERFIL_CHROME}")
    opcoes.add_argument("--profile-directory=Default")
    opcoes.add_argument("--start-maximized")
    opcoes.add_experimental_option("detach", MANTER_NAVEGADOR_ABERTO)

    try:
        driver = webdriver.Chrome(options=opcoes)
        mostrar_mensagem("[3/6] Abrindo a pasta de trabalho no Excel Online...")
        driver.get(URL_EXCEL)
    except WebDriverException as exc:
        raise RuntimeError(
            "Não foi possível iniciar o Chrome com Selenium. "
            "Feche outras janelas que estejam usando o perfil "
            f"'{PASTA_PERFIL_CHROME}' e tente novamente. Detalhe: {exc}"
        ) from exc

    mostrar_mensagem(
        "Se a Microsoft solicitar, faça login na janela do Chrome. "
        "O script não coleta nem armazena sua senha."
    )
    try:
        WebDriverWait(driver, TEMPO_PARA_LOGIN).until(
            lambda navegador: navegador.find_elements(
                By.CSS_SELECTOR,
                '#FormulaBar-NameBox-input, input[aria-label*="Name Box"], '
                'input[aria-label*="Caixa de Nome"]',
            )
            and navegador.find_elements(
                By.CSS_SELECTOR,
                '#FormulaBar-FormulaInput, input[aria-label*="Formula Bar"], '
                'textarea[aria-label*="Formula Bar"], '
                'input[aria-label*="Barra de fórmulas"], '
                'textarea[aria-label*="Barra de fórmulas"]',
            )
        )
    except TimeoutException as exc:
        raise TimeoutError(
            "A pasta de trabalho não abriu no Excel Online dentro do tempo esperado. "
            "Conclua o login, confirme a permissão de edição e verifique se URL_EXCEL "
            "é um link válido de uma pasta de trabalho."
        ) from exc

    planilha = ExcelOnlineSheet(driver, NOME_ABA)
    planilha.selecionar_aba()
    mostrar_mensagem("Pasta de trabalho do Excel Online aberta no navegador.")
    return driver, planilha


class ExcelOnlineSheet:
    """Controles de leitura e escrita da interface web do Excel Online."""

    def __init__(self, driver, nome_aba):
        self.driver = driver
        self.nome_aba = nome_aba
        self.wait = WebDriverWait(driver, 15)

    def encontrar_caixa_nome(self):
        return self.wait.until(
            lambda navegador: navegador.find_element(
                By.CSS_SELECTOR,
                '#FormulaBar-NameBox-input, input[aria-label*="Name Box"], '
                'input[aria-label*="Caixa de Nome"]',
            )
        )

    def selecionar_aba(self):
        """Seleciona a guia da planilha pelo nome visível."""
        self.driver.switch_to.default_content()
        guias = self.driver.find_elements(By.CSS_SELECTOR, '[role="tab"]')
        for guia in guias:
            texto = (guia.text or guia.get_attribute("aria-label") or "").strip()
            if texto == self.nome_aba:
                guia.click()
                time.sleep(0.5)
                return

        nomes = [
            (guia.text or guia.get_attribute("aria-label") or "").strip()
            for guia in guias
        ]
        raise ValueError(
            f"A guia '{self.nome_aba}' não foi encontrada. "
            f"Guias visíveis: {', '.join(nome for nome in nomes if nome) or 'nenhuma'}."
        )

    def selecionar_intervalo(self, intervalo):
        caixa_nome = self.encontrar_caixa_nome()
        caixa_nome.click()
        caixa_nome.send_keys(Keys.CONTROL, "a")
        caixa_nome.send_keys(intervalo)
        caixa_nome.send_keys(Keys.ENTER)
        time.sleep(0.1)

    def escrever_celula(self, endereco, valor):
        self.selecionar_intervalo(endereco)
        texto = str(valor)
        if texto.startswith(("=", "+", "-", "@")):
            texto = "'" + texto
        self.driver.switch_to.active_element.send_keys(texto)
        self.driver.switch_to.active_element.send_keys(Keys.ENTER)
        time.sleep(TEMPO_VISUAL)

    def ler_celula(self, endereco):
        self.selecionar_intervalo(endereco)
        formula = self.wait.until(
            lambda navegador: navegador.find_element(
                By.CSS_SELECTOR,
                '#FormulaBar-FormulaInput, input[aria-label*="Formula Bar"], '
                'textarea[aria-label*="Formula Bar"], '
                'input[aria-label*="Barra de fórmulas"], '
                'textarea[aria-label*="Barra de fórmulas"]',
            )
        )
        time.sleep(0.1)
        valor = formula.get_attribute("value")
        if valor is None:
            valor = formula.text
        return valor or ""

    def ler_intervalo(self, endereco):
        correspondencia = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", endereco)
        if not correspondencia:
            raise ValueError(f"Intervalo não reconhecido: {endereco}")
        coluna_inicio, linha_inicio, coluna_fim, linha_fim = correspondencia.groups()
        if linha_inicio != linha_fim:
            raise ValueError("A leitura de intervalo suporta apenas uma linha por vez.")

        colunas = self.colunas_entre(coluna_inicio, coluna_fim)
        return [
            self.ler_celula(f"{coluna}{linha_inicio}")
            for coluna in colunas
        ]

    @staticmethod
    def colunas_entre(inicio, fim):
        def numero_coluna(coluna):
            numero = 0
            for caractere in coluna:
                numero = numero * 26 + ord(caractere) - ord("A") + 1
            return numero

        def nome_coluna(numero):
            resultado = ""
            while numero:
                numero, resto = divmod(numero - 1, 26)
                resultado = chr(ord("A") + resto) + resultado
            return resultado

        primeiro = numero_coluna(inicio)
        ultimo = numero_coluna(fim)
        return [nome_coluna(numero) for numero in range(primeiro, ultimo + 1)]

    def ultima_linha_utilizada(self):
        self.selecionar_intervalo("A1")
        self.driver.switch_to.active_element.send_keys(Keys.CONTROL, Keys.ARROW_DOWN)
        time.sleep(0.15)
        endereco = self.encontrar_caixa_nome().get_attribute("value") or ""
        correspondencia = re.fullmatch(r"A(\d+)", endereco.upper())
        if not correspondencia:
            raise RuntimeError(
                f"Não foi possível identificar a linha selecionada no Excel Online: {endereco!r}"
            )

        linha = int(correspondencia.group(1))
        if linha > 1 and not self.ler_celula(f"A{linha}"):
            self.driver.switch_to.active_element.send_keys(Keys.CONTROL, Keys.ARROW_UP)
            time.sleep(0.15)
            endereco = self.encontrar_caixa_nome().get_attribute("value") or ""
            correspondencia = re.fullmatch(r"A(\d+)", endereco.upper())
            if not correspondencia:
                raise RuntimeError("Não foi possível localizar a última linha da coluna A.")
            linha = int(correspondencia.group(1))

        return max(linha, 1)

    def aplicar_cor(self, intervalo, rgb):
        """Aplica uma cor personalizada ao intervalo selecionado."""
        self.selecionar_intervalo(intervalo)
        botao = self.wait.until(
            lambda navegador: navegador.find_element(
                By.CSS_SELECTOR,
                'button[aria-label*="Fill Color"], [aria-label*="Fill Color"], '
                'button[aria-label*="Cor de Preenchimento"], '
                '[aria-label*="Cor de Preenchimento"]',
            )
        )
        botao.click()

        botao_personalizado = self.wait.until(
            lambda navegador: next(
                (
                    elemento
                    for elemento in navegador.find_elements(
                        By.CSS_SELECTOR,
                        '[role="menu"] [role="menuitem"], [role="menu"] button, '
                        '[role="dialog"] [role="button"]',
                    )
                    if "custom" in (
                        (elemento.text or "") + " " + (elemento.get_attribute("aria-label") or "")
                    ).lower()
                    or "mais cores" in (
                        (elemento.text or "") + " " + (elemento.get_attribute("aria-label") or "")
                    ).lower()
                    or "personaliz" in (
                        (elemento.text or "") + " " + (elemento.get_attribute("aria-label") or "")
                    ).lower()
                    or "more colors" in (
                        (elemento.text or "") + " " + (elemento.get_attribute("aria-label") or "")
                    ).lower()
                ),
                False,
            )
        )
        botao_personalizado.click()

        hexadecimal = "#{:02X}{:02X}{:02X}".format(*rgb)
        campo_hex = self.wait.until(
            lambda navegador: navegador.find_element(
                By.CSS_SELECTOR,
                'input[aria-label*="Hex"], input[placeholder*="Hex"], '
                'input[aria-label*="hex"], input[placeholder*="hex"], '
                'input[aria-label*="Hexadecimal"], input[placeholder*="Hexadecimal"], '
                'input[type="color"]',
            )
        )
        campo_hex.click()
        campo_hex.send_keys(Keys.CONTROL, "a")
        campo_hex.send_keys(hexadecimal)

        botoes_aplicar = self.driver.find_elements(
            By.XPATH,
            '//div[@role="dialog"]//button[normalize-space()="Apply" or normalize-space()="Aplicar" '
            'or normalize-space()="OK" or normalize-space()="Done" or normalize-space()="Concluído"]',
        )
        if not botoes_aplicar:
            raise RuntimeError("Não foi encontrado o botão para aplicar a cor personalizada.")
        botoes_aplicar[-1].click()
        time.sleep(TEMPO_VISUAL)

    def salvar(self):
        """Aguarda o salvamento automático do Excel Online."""
        if SALVAR_AUTOMATICAMENTE:
            time.sleep(1)
            mostrar_mensagem("Alterações enviadas ao Excel Online; salvamento automático ativo.")

    def fechar(self):
        if not MANTER_NAVEGADOR_ABERTO:
            self.driver.quit()


def carregar_historico():
    caminho = Path(__file__).resolve().with_name("tarefas_historico.json")
    if not caminho.exists():
        return {"processados": []}

    try:
        with caminho.open("r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        if isinstance(dados, dict) and isinstance(dados.get("processados"), list):
            return dados
        raise ValueError("Formato de histórico inválido.")
    except (json.JSONDecodeError, OSError, ValueError) as exc:
        raise RuntimeError(f"Não foi possível ler o histórico {caminho}: {exc}") from exc


def salvar_historico(historico):
    caminho = Path(__file__).resolve().with_name("tarefas_historico.json")
    with caminho.open("w", encoding="utf-8") as arquivo:
        json.dump(historico, arquivo, ensure_ascii=False, indent=2)


def gerar_id_tarefa(tarefa):
    identidade = tarefa.get("id")
    if identidade:
        return str(identidade).strip()
    return gerar_id_compatibilidade(tarefa)


def gerar_id_compatibilidade(tarefa):
    payload = {
        "tarefa": str(tarefa.get("tarefa", "")).strip().lower(),
        "tempo_medio": str(tarefa.get("tempo_medio", "")).strip().lower(),
        "descricao": str(tarefa.get("descricao", "")).strip().lower(),
        "status": str(tarefa.get("status", "")).strip().lower(),
        "ajuda_externa": str(tarefa.get("ajuda_externa", "")).strip().lower(),
        "urgencia": int(tarefa.get("urgencia", 1)),
    }
    return hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()


def tarefa_ja_processada(tarefa, historico):
    processados = set(historico["processados"])
    return (
        gerar_id_tarefa(tarefa) in processados
        or gerar_id_compatibilidade(tarefa) in processados
    )


def marcar_tarefa_processada(tarefa, historico):
    identificadores = {
        gerar_id_tarefa(tarefa),
        gerar_id_compatibilidade(tarefa),
    }
    novos_ids = identificadores.difference(historico["processados"])
    if novos_ids:
        historico["processados"].extend(sorted(novos_ids))
        salvar_historico(historico)


def validar_status(status):
    if isinstance(status, int) or str(status).strip().isdigit():
        numero = int(status)
        if numero in STATUS_TAREFA:
            return STATUS_TAREFA[numero]
        raise ValueError("Status inválido. Use um número de 1 a 5.")

    texto = str(status).strip()
    for valor in STATUS_TAREFA.values():
        if valor.lower() == texto.lower():
            return valor
    raise ValueError("Status inválido. Use um número de 1 a 5 ou um status permitido.")


def validar_urgencia(urgencia):
    try:
        valor = int(urgencia)
    except (TypeError, ValueError) as exc:
        raise ValueError("Urgência deve ser um número inteiro de 1 a 5.") from exc
    if valor not in CORES_URGENCIA:
        raise ValueError("Urgência deve estar entre 1 e 5.")
    return valor


def validar_realizacoes(realizacoes):
    try:
        valor = int(realizacoes)
    except (TypeError, ValueError) as exc:
        raise ValueError("Realizações deve ser um número inteiro não negativo.") from exc
    if valor < 0:
        raise ValueError("Realizações não podem ser negativas.")
    return valor


def normalizar_tarefa(tarefa):
    if not isinstance(tarefa, dict):
        raise ValueError("Cada tarefa deve ser um dicionário.")
    faltando = [campo for campo in CAMPOS_TAREFA if campo not in tarefa]
    if faltando:
        raise ValueError(f"Tarefa incompleta. Campos ausentes: {', '.join(faltando)}.")

    limpa = {
        "tarefa": str(tarefa["tarefa"]).strip(),
        "tempo_medio": str(tarefa["tempo_medio"]).strip(),
        "descricao": str(tarefa["descricao"]).strip(),
        "realizacoes": validar_realizacoes(tarefa["realizacoes"]),
        "status": validar_status(tarefa["status"]),
        "ajuda_externa": str(tarefa["ajuda_externa"]).strip(),
        "urgencia": validar_urgencia(tarefa["urgencia"]),
    }
    for campo in ("tarefa", "tempo_medio", "descricao", "ajuda_externa"):
        if not limpa[campo]:
            raise ValueError(f"O campo '{campo}' não pode ficar vazio.")
    if tarefa.get("id"):
        limpa["id"] = str(tarefa["id"]).strip()
    return limpa


def obter_ultima_linha_utilizada(ws):
    return ws.ultima_linha_utilizada()


def linha_disponivel(ws):
    return max(obter_ultima_linha_utilizada(ws) + 1, 2)


def salvar_arquivo(wb, mensagem_sucesso="Planilha salva."):
    wb.salvar()
    mostrar_mensagem(mensagem_sucesso)


def aplicar_cor_urgencia(ws, linha, urgencia):
    ws.aplicar_cor(f"A{linha}:G{linha}", CORES_URGENCIA[int(urgencia)])


def cadastrar_tarefa_no_excel(ws, tarefa, historico=None, silencioso=False):
    tarefa_validada = normalizar_tarefa(tarefa)
    linha = linha_disponivel(ws)
    valores = [
        tarefa_validada["tarefa"],
        tarefa_validada["tempo_medio"],
        tarefa_validada["descricao"],
        tarefa_validada["realizacoes"],
        tarefa_validada["status"],
        tarefa_validada["ajuda_externa"],
        tarefa_validada["urgencia"],
    ]

    for coluna, valor in zip("ABCDEFG", valores):
        endereco = f"{coluna}{linha}"
        mostrar_mensagem(f"Preenchendo célula {endereco}...")
        ws.escrever_celula(endereco, valor)

    mostrar_mensagem("Aplicando cor de urgência...")
    aplicar_cor_urgencia(ws, linha, tarefa_validada["urgencia"])
    salvar_arquivo(ws, "Tarefa enviada ao Excel Online.")
    if historico is not None:
        marcar_tarefa_processada(tarefa_validada, historico)
    if not silencioso:
        mostrar_mensagem(f"Tarefa cadastrada na linha {linha}.")
    time.sleep(DELAY_TAREFA)
    return linha


def processar_lista_automatica(ws, historico, tarefas):
    if not tarefas:
        mostrar_mensagem("Nenhuma tarefa na lista automática.")
        return

    mostrar_mensagem("[4/6] Identificando próxima linha disponível...")
    mostrar_mensagem("[5/6] Validando tarefas...")
    tarefas_validadas = [normalizar_tarefa(tarefa) for tarefa in tarefas]
    mostrar_mensagem("[6/6] Iniciando preenchimento visual...")

    for indice, tarefa in enumerate(tarefas_validadas, start=1):
        if tarefa_ja_processada(tarefa, historico):
            mostrar_mensagem(f"Tarefa '{tarefa['tarefa']}' já processada; ignorada.")
            continue
        resposta = input(
            f"A tarefa '{tarefa['tarefa']}' ainda não foi cadastrada. Incluir? [S/N]: "
        ).strip().lower()
        if resposta != "s":
            mostrar_mensagem("Tarefa ignorada.")
            continue
        mostrar_mensagem(f"\nTarefa {indice} de {len(tarefas_validadas)}: {tarefa['tarefa']}")
        cadastrar_tarefa_no_excel(ws, tarefa, historico=historico)
        time.sleep(DELAY_FINAL)


def obter_status_por_numero(mensagem):
    while True:
        try:
            return validar_status(input(mensagem).strip())
        except ValueError as exc:
            mostrar_mensagem(f"Valor inválido: {exc}")


def obter_urgencia_por_numero(mensagem):
    while True:
        try:
            return validar_urgencia(input(mensagem).strip())
        except ValueError as exc:
            mostrar_mensagem(f"Valor inválido: {exc}")


def validar_aba(ws):
    ws.selecionar_aba()
    return ws


def visualizar_tarefas(ws):
    ultima_linha = obter_ultima_linha_utilizada(ws)
    if ultima_linha <= 1:
        mostrar_mensagem("Nenhuma tarefa cadastrada.")
        return

    mostrar_mensagem("\n--- Tarefas cadastradas ---")
    for linha in range(2, ultima_linha + 1):
        valores = ws.ler_intervalo(f"A{linha}:G{linha}")
        if not any(str(valor).strip() for valor in valores):
            continue
        mostrar_mensagem(
            f"Linha {linha}: {valores[0]} | Tempo: {valores[1]} | Descrição: {valores[2]} "
            f"| Realizações: {valores[3]} | Status: {valores[4]} | Ajuda: {valores[5]} "
            f"| Urgência: {valores[6]}"
        )


def solicitar_linha(ws, mensagem):
    ultima_linha = obter_ultima_linha_utilizada(ws)
    if ultima_linha <= 1:
        mostrar_mensagem("Nenhuma tarefa cadastrada.")
        return None
    try:
        linha = int(input(mensagem).strip())
    except ValueError:
        mostrar_mensagem("Informe um número de linha válido.")
        return None
    if linha < 2 or linha > ultima_linha:
        mostrar_mensagem("Linha inválida.")
        return None
    return linha


def atualizar_status_tarefa(ws):
    linha = solicitar_linha(ws, "Linha da tarefa: ")
    if linha is None:
        return
    status = obter_status_por_numero("Novo status (1 a 5): ")
    ws.escrever_celula(f"E{linha}", status)
    salvar_arquivo(ws, "Status atualizado.")


def alterar_urgencia_tarefa(ws):
    linha = solicitar_linha(ws, "Linha da tarefa: ")
    if linha is None:
        return
    urgencia = obter_urgencia_por_numero("Nova urgência (1 a 5): ")
    ws.escrever_celula(f"G{linha}", urgencia)
    aplicar_cor_urgencia(ws, linha, urgencia)
    salvar_arquivo(ws, "Urgência e cor atualizadas.")


def registrar_nova_realizacao(ws):
    linha = solicitar_linha(ws, "Linha da tarefa: ")
    if linha is None:
        return
    atual = ws.ler_celula(f"D{linha}") or "0"
    try:
        incremento = validar_realizacoes(input("Quantas realizações registrar? ").strip())
        total = validar_realizacoes(int(float(atual)) + incremento)
    except (ValueError, TypeError) as exc:
        mostrar_mensagem(f"Valor inválido: {exc}")
        return
    ws.escrever_celula(f"D{linha}", total)
    salvar_arquivo(ws, "Realizações atualizadas.")


def editar_tarefa(ws):
    linha = solicitar_linha(ws, "Linha da tarefa: ")
    if linha is None:
        return
    mostrar_mensagem(
        "Campos: 1=Tarefa, 2=Tempo, 3=Descrição, 4=Realizações, "
        "5=Status, 6=Ajuda externa, 7=Urgência"
    )
    try:
        campo = int(input("Campo: ").strip())
    except ValueError:
        mostrar_mensagem("Campo inválido.")
        return

    colunas = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G"}
    if campo not in colunas:
        mostrar_mensagem("Campo inválido.")
        return
    if campo == 4:
        valor = validar_realizacoes(input("Nova quantidade de realizações: ").strip())
    elif campo == 5:
        valor = obter_status_por_numero("Novo status (1 a 5): ")
    elif campo == 7:
        valor = obter_urgencia_por_numero("Nova urgência (1 a 5): ")
    else:
        valor = input("Novo valor: ").strip()
        if not valor:
            mostrar_mensagem("O valor não pode ficar vazio.")
            return

    ws.escrever_celula(f"{colunas[campo]}{linha}", valor)
    if campo == 7:
        aplicar_cor_urgencia(ws, linha, valor)
    salvar_arquivo(ws, "Edição salva.")


def cadastrar_tarefa_manual(ws, historico):
    tarefa = {
        "tarefa": input("Nome da tarefa: ").strip(),
        "tempo_medio": input("Tempo médio: ").strip(),
        "descricao": input("Descrição: ").strip(),
        "realizacoes": validar_realizacoes(input("Realizações: ").strip()),
        "status": obter_status_por_numero("Status (1 a 5): "),
        "ajuda_externa": input("Ajuda externa (Sim/Não): ").strip(),
        "urgencia": obter_urgencia_por_numero("Urgência (1 a 5): "),
    }
    cadastrar_tarefa_no_excel(ws, tarefa, historico=historico)


def menu_principal(ws, historico):
    while True:
        print("\n=========================================")
        print("       SISTEMA DE CONTROLE DE TAREFAS")
        print("=========================================")
        print("1 - Cadastrar nova tarefa")
        print("2 - Visualizar tarefas cadastradas")
        print("3 - Atualizar status de uma tarefa")
        print("4 - Alterar grau de urgência")
        print("5 - Registrar nova realização")
        print("6 - Editar informações de uma tarefa")
        print("7 - Finalizar programa")
        print("=========================================")
        opcao = input("Escolha uma opção: ").strip()

        if opcao == "1":
            cadastrar_tarefa_manual(ws, historico)
        elif opcao == "2":
            visualizar_tarefas(ws)
        elif opcao == "3":
            atualizar_status_tarefa(ws)
        elif opcao == "4":
            alterar_urgencia_tarefa(ws)
        elif opcao == "5":
            registrar_nova_realizacao(ws)
        elif opcao == "6":
            editar_tarefa(ws)
        elif opcao == "7":
            salvar_arquivo(ws, "Excel Online salva automaticamente.")
            return
        else:
            mostrar_mensagem("Opção inválida.")
        input("\nPressione Enter para continuar...")


def finalizar_automacao(driver, ws):
    salvar_arquivo(ws, "Finalizado; alterações sincronizadas pelo Excel Online.")
    if MANTER_NAVEGADOR_ABERTO:
        mostrar_mensagem("Chrome mantido aberto.")
    else:
        driver.quit()


def main():
    driver = None
    try:
        validar_configuracoes()
        driver, ws = abrir_excel_online()
        validar_aba(ws)
        historico = carregar_historico()
        processar_lista_automatica(ws, historico, TAREFAS_PARA_CADASTRAR)
        salvar_arquivo(ws, "Lote automático concluído.")
        menu_principal(ws, historico)
        finalizar_automacao(driver, ws)
    except Exception as exc:
        mostrar_mensagem(f"\nErro durante a automação: {exc}")
        if driver is not None and not MANTER_NAVEGADOR_ABERTO:
            driver.quit()
        input("Pressione Enter para encerrar...")
        raise


if __name__ == "__main__":
    main()
