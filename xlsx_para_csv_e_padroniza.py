from pathlib import Path
import re
import unicodedata
import pandas as pd

PASTA_ENTRADA = r"C:\Users\Filippo Di Giuseppe\Downloads\arquivosConverter"
PASTA_SAIDA = r"C:\Users\Filippo Di Giuseppe\Downloads\arquivosConvertido"
SEPARADOR_CSV = ","
ENCODING = "utf-8-sig"

COLUNAS_A_REMOVER_NO_FINAL = ["idp_original", "familia_arquivo", "ordem_arquivo"]


def normalizar_texto(valor: str) -> str:
    valor = str(valor or "").strip().lower()
    valor = unicodedata.normalize("NFD", valor)
    valor = "".join(ch for ch in valor if unicodedata.category(ch) != "Mn")
    valor = re.sub(r"[^\w\s]+", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def nome_coluna_padrao(coluna: str) -> str:
    base = normalizar_texto(coluna).replace(" ", "_")
    mapa = {
        "idp": "idp",
        "estado": "estado",
        "cidade": "cidade",
        "cod_estado": "cod_estado",
        "cod_municipio": "cod_municipio",
        "cod_meso": "cod_meso",
        "cod_micro": "cod_micro",
        "regiao_cidade": "regiao_cidade",
        "setor_dentro_da_cidade": "setor_dentro_cidade",
        "setor_dentro_cidade": "setor_dentro_cidade",
        "sexo": "sexo",
        "idade": "idade",
        "grau_de_instrucao": "grau_instrucao",
        "grau_instrucao": "grau_instrucao",
        "ocupacao": "ocupacao",
        "renda_familiar": "renda_familiar",
        "data_da_pesquisa": "data_pesquisa",
        "categoria": "categoria",
        "nome": "nome",
        "endereco_completo": "endereco_completo",
        "endereco": "endereco",
        "numero": "numero",
        "bairro": "bairro",
        "cep": "cep",
        "email": "email",
        "tf1": "tf1",
        "tf2": "tf2",
        "tf3": "tf3",
        "tf4": "tf4",
        "n_pesq": "numero_pesq",
        "numero_pesq": "numero_pesq",
        "dia_pesq": "dia_pesq",
        "qt_tf1": "qt_tf1",
        "qt_tf2": "qt_tf2",
        "qt_tf3": "qt_tf3",
        "qt_tf4": "qt_tf4",
    }
    return mapa.get(base, base)


def limpar_telefone(valor) -> str:
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    if s.lower() in {"nan", "none", ""}:
        return ""
    return re.sub(r"\D", "", s)


def linha_tem_conteudo(linha) -> bool:
    for valor in linha:
        if pd.isna(valor):
            continue
        if str(valor).strip() != "":
            return True
    return False


def normalizar_sexo(valor: str) -> str:
    s = str(valor or "").strip().upper()
    mapa = {
        "M": "M",
        "F": "F",
        "MASC": "M",
        "FEM": "F",
        "MASCULINO": "M",
        "FEMININO": "F",
    }
    return mapa.get(s, "")


def normalizar_estado(valor: str) -> str:
    s = str(valor or "").strip().upper()
    s = re.sub(r"[^A-Z]", "", s)
    return s[:2]


def obter_familia_arquivo(nome_arquivo: str) -> str:
    nome_base = Path(nome_arquivo).stem.strip()
    familia = re.sub(r"(?i)\bTOTAL\s+(\d+)\b", "TOTAL", nome_base)
    familia = re.sub(r"\s+", " ", familia).strip()
    return familia


def obter_ordem_arquivo(nome_arquivo: str) -> int:
    nome_base = Path(nome_arquivo).stem.strip()
    match = re.search(r"(?i)\bTOTAL\s+(\d+)\b", nome_base)
    if match:
        return int(match.group(1))
    return 1


def escolher_melhor_aba(caminho_xlsx: Path) -> str:
    excel_file = pd.ExcelFile(caminho_xlsx)
    abas = excel_file.sheet_names

    if not abas:
        raise ValueError("Nenhuma aba encontrada no arquivo.")

    melhor_aba = None
    melhor_pontuacao = -1

    for aba in abas:
        pontuacao = 0
        nome = normalizar_texto(aba)

        for termo in ["com", "priv", "fac", "pesq", "telefone", "base", "dados", "planilha"]:
            if termo in nome:
                pontuacao += 1

        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_aba = aba

    return melhor_aba or abas[0]


def detectar_header_linha(df_sem_header: pd.DataFrame) -> int:
    melhores_colunas = {
        "idp", "estado", "cidade", "sexo", "idade",
        "tf1", "tf2", "tf3", "tf4", "nome",
        "bairro", "email", "cep", "endereco", "categoria"
    }

    melhor_idx = 0
    melhor_score = -1
    limite = min(len(df_sem_header), 15)

    for idx in range(limite):
        linha = df_sem_header.iloc[idx].tolist()
        normalizados = {nome_coluna_padrao(str(c)) for c in linha if str(c).strip() != ""}
        score = len(normalizados.intersection(melhores_colunas))

        if score > melhor_score:
            melhor_score = score
            melhor_idx = idx

    return melhor_idx


def ler_planilha_flexivel(caminho_xlsx: Path) -> pd.DataFrame:
    aba_escolhida = escolher_melhor_aba(caminho_xlsx)
    print(f"  -> Aba usada: {aba_escolhida}")

    df_bruto = pd.read_excel(
        caminho_xlsx,
        sheet_name=aba_escolhida,
        header=None,
        dtype=object
    )

    if df_bruto.empty:
        return pd.DataFrame()

    header_idx = detectar_header_linha(df_bruto)

    df = pd.read_excel(
        caminho_xlsx,
        sheet_name=aba_escolhida,
        header=header_idx,
        dtype=object
    )

    return df


def ler_e_limpar_arquivo(caminho_xlsx: Path) -> pd.DataFrame:
    print(f"Processando: {caminho_xlsx.name}")

    df = ler_planilha_flexivel(caminho_xlsx)

    if df.empty:
        print("  -> Arquivo vazio após leitura.")
        return df

    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.isna()]
    df.columns = [nome_coluna_padrao(c) for c in df.columns]

    df = df.dropna(how="all")
    df = df[df.apply(linha_tem_conteudo, axis=1)].copy()

    if "idp" in df.columns:
        df["idp_original"] = df["idp"]
    else:
        df["idp_original"] = ""

    if "sexo" in df.columns:
        df["sexo"] = df["sexo"].apply(normalizar_sexo)
        df = df[df["sexo"].isin(["M", "F"])].copy()
    else:
        print("  -> Coluna sexo não encontrada. Arquivo ficará sem linhas.")
        df = df.iloc[0:0].copy()

    if "estado" in df.columns:
        df["estado"] = df["estado"].apply(normalizar_estado)
    else:
        df["estado"] = ""

    df["arquivo_origem"] = caminho_xlsx.name
    df["familia_arquivo"] = obter_familia_arquivo(caminho_xlsx.name)
    df["ordem_arquivo"] = obter_ordem_arquivo(caminho_xlsx.name)

    for col in ["tf1", "tf2", "tf3", "tf4"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].apply(limpar_telefone)

    print(f"  -> Linhas válidas após filtro de sexo: {len(df)}")
    return df


def reatribuir_idp_por_familia(dataframes_por_familia: dict):
    resultado = []

    for familia, itens in sorted(dataframes_por_familia.items(), key=lambda x: x[0].lower()):
        print(f"\nFamília: {familia}")

        itens_ordenados = sorted(itens, key=lambda x: (x[0], x[1].lower()))
        proximo_idp = 1

        for ordem, nome_arquivo, df in itens_ordenados:
            qtd = len(df)

            if qtd == 0:
                print(f"  -> {nome_arquivo}: 0 linha(s)")
                continue

            df = df.copy()
            df["idp"] = range(proximo_idp, proximo_idp + qtd)

            inicio = proximo_idp
            fim = proximo_idp + qtd - 1

            proximo_idp += qtd
            resultado.append((nome_arquivo, df))

            print(f"  -> {nome_arquivo}: {qtd} linha(s) | IDP {inicio} até {fim}")

    return resultado


def preparar_df_para_salvar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df = df.drop(columns=COLUNAS_A_REMOVER_NO_FINAL, errors="ignore")

    # garante que idp fique como primeira coluna, se existir
    colunas = list(df.columns)
    if "idp" in colunas:
        colunas.remove("idp")
        colunas = ["idp"] + colunas
        df = df[colunas]

    return df


def salvar_csv(caminho_saida: Path, df: pd.DataFrame):
    df_final = preparar_df_para_salvar(df)
    df_final.to_csv(caminho_saida, index=False, sep=SEPARADOR_CSV, encoding=ENCODING)


def main():
    pasta_entrada = Path(PASTA_ENTRADA)
    pasta_saida = Path(PASTA_SAIDA)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    arquivos = sorted(pasta_entrada.glob("*.xlsx"))

    if not arquivos:
        print("Nenhum arquivo .xlsx encontrado.")
        return

    print(f"Total de arquivos encontrados: {len(arquivos)}")

    dataframes_por_familia = {}

    for caminho_xlsx in arquivos:
        try:
            df = ler_e_limpar_arquivo(caminho_xlsx)

            familia = obter_familia_arquivo(caminho_xlsx.name)
            ordem = obter_ordem_arquivo(caminho_xlsx.name)

            dataframes_por_familia.setdefault(familia, []).append(
                (ordem, caminho_xlsx.name, df)
            )

        except Exception as e:
            print(f"ERRO em {caminho_xlsx.name}: {e}")

    arquivos_processados = reatribuir_idp_por_familia(dataframes_por_familia)

    print("\nSalvando CSVs...")
    for nome_arquivo, df in arquivos_processados:
        try:
            caminho_saida = pasta_saida / Path(nome_arquivo).with_suffix(".csv").name
            salvar_csv(caminho_saida, df)
            print(f"  -> CSV salvo em: {caminho_saida} | Linhas finais: {len(df)}")
        except Exception as e:
            print(f"ERRO ao salvar {nome_arquivo}: {e}")

    print("\nConcluído.")


if __name__ == "__main__":
    main()