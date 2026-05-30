import { TARGET_COLUMNS, ColumnMappingConfig, ExcelRow, TransformationType } from './types';

// Regular expressions to assist in smart detection
const CORRESPONDENCIA_DICIONARIO: { [key: string]: string[] } = {
  id_da_venda: ['venda', 'id', 'num', 'numero', 'codigo', 'cod_venda', 'order_id', 'id_da_venda', 'transação', 'id_compra', 'order'],
  data_do_pagamento: ['data', 'pagamento', 'pago', 'data_do_pagamento', 'created_at', 'pago_em', 'date', 'data_compra', 'data_pagto', 'data_venda'],
  id_do_produto: ['produto_id', 'id_prod', 'cod_produto', 'id_do_produto', 'prod_id', 'idprod'],
  nome_do_produto: ['produto', 'nome_produto', 'item', 'descrição', 'descricao', 'nome_do_produto', 'curso_comprado', 'product'],
  id_do_sku: ['sku_id', 'cod_sku', 'sku_code', 'id_do_sku', 'sku', 'ref_sku', 'ref'],
  nome_do_sku: ['nome_sku', 'variação', 'variacao', 'nome_do_sku', 'variacao_sku'],
  valor_unitario: ['valor', 'preco', 'preço', 'valor_unitario', 'price', 'unit_p', 'valor_paco', 'unitario'],
  moeda: ['moeda', 'currency'],
  quantidade: ['quantidade', 'quant', 'qtd', 'quantity', 'unidades', 'quantidade_itens', 'itens'],
  nome_do_cliente: ['cliente', 'nome', 'comprador', 'nome_do_cliente', 'client', 'nome_completo'],
  email_do_cliente: ['email', 'e-mail', 'email_do_cliente', 'mail', 'email_contato'],
  telefone_do_cliente: ['telefone', 'tel', 'celular', 'cel', 'phone', 'ddd_telefone', 'contato', 'apenas_celular'],
  tipo_de_cliente: ['tipo', 'tipo_cliente', 'pessoa', 'tipo_de_cliente', 'tipo_pessoa'],
  documento_do_cliente: ['documento', 'cpf', 'cnpj', 'doc', 'cpf_cnpj', 'documento_do_cliente', 'cpf_cnpj_cliente', 'documento_identificação', 'cpf_cnpj_comprador'],
  endereco_rua: ['rua', 'logradouro', 'endereco', 'endereço', 'endereco_rua', 'street', 'rua_entrega'],
  endereco_numero: ['numero', 'num', 'endereco_numero', 'number', 'numero_casa', 'nro'],
  endereco_bairro: ['bairro', 'endereco_bairro', 'district'],
  endereco_complemento: ['complemento', 'comp', 'endereco_complemento', 'complemento_apto'],
  endereco_cidade: ['cidade', 'municipio', 'endereco_cidade', 'city', 'cidade_cadastro'],
  endereco_estado: ['estado', 'uf', 'endereco_estado', 'state', 'estado_cadastro'],
  endereco_cep: ['cep', 'zip', 'zipcode', 'endereco_cep', 'cep_entrega'],
  endereco_pais: ['pais', 'país', 'endereco_pais', 'country'],
};

// Cleans dynamic strings for token comparison
function norm(str: string): string {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // remove accents
    .replace(/[^a-z0-9]/g, ''); // keep alphanumeric only
}

/**
 * Smart Mapping Suggestion
 * Will look at list of spreadsheet columns (keys) and find closest matches for each target layout path
 */
export function sugerirMapeamentoAutomatico(
  uploadedColumns: string[],
  targetKey: string
): { sourceType: 'column' | 'fixed'; sourceColumn: string; fixedValue: string } {
  const normalizedUploaded = uploadedColumns.map((col) => ({
    original: col,
    normalized: norm(col),
  }));

  const matches = CORRESPONDENCIA_DICIONARIO[targetKey] || [];

  // 1. Try to find direct match or match based on dictionary aliases
  for (const matchAlias of matches) {
    const aliasNorm = norm(matchAlias);
    const found = normalizedUploaded.find(
      (col) => col.normalized === aliasNorm || col.normalized.includes(aliasNorm) || aliasNorm.includes(col.normalized)
    );
    if (found) {
      return { sourceType: 'column', sourceColumn: found.original, fixedValue: '' };
    }
  }

  // 2. Specific fallbacks for default values if not mapped
  const targetCol = TARGET_COLUMNS.find((tc) => tc.key === targetKey);
  if (targetCol?.defaultFixed) {
    return { sourceType: 'fixed', sourceColumn: '', fixedValue: targetCol.defaultFixed };
  }

  return { sourceType: 'column', sourceColumn: '', fixedValue: '' };
}

// Formatters
function formatDateISO(date: Date): string {
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`; // BR format is widely preferred for invoices, but support ISO too. Let's do YYYY-MM-DD for easier import or DD/MM/YYYY.
}

export function parseAndConvertDate(val: any): string {
  if (val === undefined || val === null) return '';
  const str = String(val).trim();
  if (!str) return '';

  // 1. Excel serial dates (numbers like 45000)
  if (/^\d{5}(\.\d+)?$/.test(str)) {
    const serial = parseFloat(str);
    const excelEpoch = new Date(1899, 11, 30); // Dec 30 1899 is the Excel epoch start
    const date = new Date(excelEpoch.getTime() + serial * 86400000);
    if (!isNaN(date.getTime())) {
      return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
    }
  }

  // 2. Already formatted dates like DD/MM/YYYY or YYYY-MM-DD
  const brMatch = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (brMatch) {
    return `${brMatch[1].padStart(2, '0')}/${brMatch[2].padStart(2, '0')}/${brMatch[3]}`;
  }

  const isoMatch = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (isoMatch) {
    return `${isoMatch[3].padStart(2, '0')}/${isoMatch[2].padStart(2, '0')}/${isoMatch[1]}`;
  }

  // General parse fallback
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return `${String(parsed.getDate()).padStart(2, '0')}/${String(parsed.getMonth() + 1).padStart(2, '0')}/${parsed.getFullYear()}`;
  }

  return str;
}

export function formatCPF(val: string): string {
  const digits = val.replace(/\D/g, '');
  if (digits.length !== 11) return val; // Return raw if not standard
  return `${digits.substring(0, 3)}.${digits.substring(3, 6)}.${digits.substring(6, 9)}-${digits.substring(9, 11)}`;
}

export function formatCNPJ(val: string): string {
  const digits = val.replace(/\D/g, '');
  if (digits.length !== 14) return val;
  return `${digits.substring(0, 2)}.${digits.substring(2, 5)}.${digits.substring(5, 8)}/${digits.substring(8, 12)}-${digits.substring(12, 14)}`;
}

export function formatCEP(val: string): string {
  const digits = val.replace(/\D/g, '');
  if (digits.length !== 8) return val;
  return `${digits.substring(0, 5)}-${digits.substring(5, 8)}`;
}

export function applyTransformation(val: any, type: TransformationType): string {
  if (val === undefined || val === null) return '';
  let str = String(val).trim();

  switch (type) {
    case 'remover_especiais':
      // Remove characters that are not normal letters (keeps accented ones), numbers, hyphen, dot, at-sign, or space.
      return str.replace(/[^a-zA-Z0-9\sÀ-ÿ\-@.]/g, '');
    case 'remover_espacos':
      return str.trim().replace(/\s+/g, ' ');
    case 'maiusculo':
      return str.toUpperCase();
    case 'minusculo':
      return str.toLowerCase();
    case 'formatar_cpf':
      return formatCPF(str);
    case 'formatar_cnpj':
      return formatCNPJ(str);
    case 'formatar_cep':
      return formatCEP(str);
    case 'converter_data':
      return parseAndConvertDate(str);
    case 'none':
    default:
      return str;
  }
}

/**
 * Execute whole Row Transformation pipeline based on mapped configuration
 */
export function processarLinha(
  originalRow: ExcelRow,
  config: ColumnMappingConfig[],
  rowIndex: number = 0
): ExcelRow {
  const outRow: ExcelRow = {};

  for (const target of TARGET_COLUMNS) {
    const colConfig = config.find((c) => c.targetKey === target.key);
    let finalValue = '';

    if (colConfig) {
      if (colConfig.sourceType === 'column') {
        const sourceColName = colConfig.sourceColumn;
        finalValue = sourceColName ? String(originalRow[sourceColName] ?? '') : '';
      } else if (colConfig.sourceType === 'fixed') {
        finalValue = colConfig.fixedValue;
      } else if (colConfig.sourceType === 'concat') {
        const parts = colConfig.concatColumns
          .map((colName) => String(originalRow[colName] ?? ''))
          .filter((p) => p.trim() !== '');
        finalValue = parts.join(colConfig.concatSeparator || ' ');
      }

      // Execute sequential transformations
      if (colConfig.transformations && colConfig.transformations.length > 0) {
        for (const trans of colConfig.transformations) {
          finalValue = applyTransformation(finalValue, trans);
        }
      }
    } else {
      // Default fallback
      finalValue = target.defaultFixed || '';
    }

    outRow[target.key] = finalValue;
  }

  // 1. Regra para documento do cliente vazio: preenche com número de 5 dígitos que não se repete (pseudo-aleatório e estável)
  const docResult = String(outRow['documento_do_cliente'] || '').trim();
  if (docResult === '') {
    // Gerador determinístico baseado no índice com boa dispersão de 10000 a 99999
    const randomNum = 10000 + ((rowIndex * 17923 + 45329) % 90000);
    outRow['documento_do_cliente'] = String(randomNum);
  }

  // 2. Regra para tipo_de_cliente baseado em documento_do_cliente
  const finalDoc = String(outRow['documento_do_cliente'] || '').trim();
  const digits = finalDoc.replace(/\D/g, '');

  if (digits.length === 14) {
    outRow['tipo_de_cliente'] = 'pessoa_juridica';
  } else if (digits.length === 11) {
    outRow['tipo_de_cliente'] = 'pessoa_fisica';
  } else {
    outRow['tipo_de_cliente'] = 'internacional';
  }

  return outRow;
}

// SAMPLE DATA GENERATOR FOR TECH STORE CO.
export const SAMPLE_TECH_STORE_ROWS: ExcelRow[] = [
  {
    Código: 'VD-83921',
    Data_Compra: '46170.83333', // Excel serial date representing May 30, 2026 approx
    Produto_Nome: 'Teclado Mecânico Keychron K2 v2 RGB',
    Ref_SKU: 'KEY-K2-RGB-BROWN',
    Preço: '799.90',
    Qtd: '1',
    Cliente_Nome: 'Luiz Inácio Silva',
    Cliente_Email: 'luiz.silva@gmail.com',
    Apenas_Celular: '11988887777',
    CPF_CNPJ: '49929949182',
    Logradouro: 'Avenida Paulista',
    Num: '1200',
    Bairro: 'Bela Vista',
    Cidade_UF: 'São Paulo/SP', // Field to showcase parsing/concat or splitting
    CEP: '01310100',
  },
  {
    Código: 'VD-83922',
    Data_Compra: '30/05/2026',
    Produto_Nome: 'Mouse Sem Fio Logitech MX Master 3S',
    Ref_SKU: 'LOGI-MX3S-GRA',
    Preço: '649.90',
    Qtd: '2',
    Cliente_Nome: 'Mariana de Oliveira Santos',
    Cliente_Email: 'mariana.santos@outlook.com',
    Apenas_Celular: '21977776666',
    CPF_CNPJ: '01293847291',
    Logradouro: 'Rua Voluntários da Pátria',
    Num: '450',
    Bairro: 'Botafogo',
    Cidade_UF: 'Rio de Janeiro/RJ',
    CEP: '22270010',
  },
  {
    Código: 'VD-83923',
    Data_Compra: '2026-05-29',
    Produto_Nome: 'Monitor Gamer LG Ultrawide 29UM69G',
    Ref_SKU: 'LG-UW-29-GAMER',
    Preço: '1299.00',
    Qtd: '1',
    Cliente_Nome: 'Fernando Henrique Cardoso',
    Cliente_Email: 'fhc@sociologia.org',
    Apenas_Celular: '3132221111',
    CPF_CNPJ: '83920192837',
    Logradouro: 'Praça da Liberdade',
    Num: 's/n',
    Bairro: 'Funcionários',
    Cidade_UF: 'Belo Horizonte/MG',
    CEP: '30140010',
  },
  {
    Código: 'VD-83924',
    Data_Compra: '28-05-2026',
    Produto_Nome: 'Fone de Ouvido Sony WH-1000XM5 ANC',
    Ref_SKU: 'SONY-XM5-BLACK',
    Preço: '2199.00',
    Qtd: '1',
    Cliente_Nome: 'Amanda Rebeca Menezes',
    Cliente_Email: 'amanda@menezes.etc.br',
    Apenas_Celular: '81 99999-5511',
    CPF_CNPJ: '12839274811',
    Logradouro: 'Rua da Aurora',
    Num: '100',
    Bairro: 'Boa Vista',
    Cidade_UF: 'Recife/PE',
    CEP: '50050000',
  }
];

// SAMPLE DATA GENERATOR FOR WEB ACADEMY
export const SAMPLE_WEB_ACADEMY_ROWS: ExcelRow[] = [
  {
    id_compra: '99418',
    data_pagto: '2026-05-30',
    curso_comprado: 'MBA em Arquitetura de Software Core & Cloud',
    cod_sku: 'MBA-ARCH-CLOUD',
    valor: '4500.00',
    quantidade_itens: '1',
    nome_completo: 'Carlos Eduardo',
    sobrenome: 'Melo Ramos',
    email_contato: 'carlos.meloramos@edu.br',
    telefone_ddd: '(11) 94444-2211',
    documento_identificação: '22333444000199', // CNPJ format showcase
    rua_entrega: 'Alameda Santos',
    numero_casa: '2101',
    complemento_apto: 'Bl C Apto 112',
    cidade_cadastro: 'São Paulo',
    estado_cadastro: 'SP',
    cep_entrega: '01419-002',
  },
  {
    id_compra: '99419',
    data_pagto: '29/05/2026',
    curso_comprado: 'Formação Avançada de Engenharia e GenAI',
    cod_sku: 'DEV-GEN-AI-ADV',
    valor: '1890.00',
    quantidade_itens: '1',
    nome_completo: 'Sofia de',
    sobrenome: 'Medeiros Carvalho',
    email_contato: 'sofia.carvalho@empresa.com.br',
    telefone_ddd: '(31) 98888-9999',
    documento_identificação: '88392183921',
    rua_entrega: 'Avenida Getúlio Vargas',
    numero_casa: '800',
    complemento_apto: 'Prédio A, Sala 304',
    cidade_cadastro: 'Belo Horizonte',
    estado_cadastro: 'MG',
    cep_entrega: '30112-020',
  }
];
