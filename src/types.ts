export interface TargetColumnDefinition {
  key: string;
  label: string;
  description: string;
  placeholder?: string;
  defaultFixed?: string;
}

export type SourceType = 'column' | 'fixed' | 'concat';

export type TransformationType =
  | 'none'
  | 'remover_especiais'
  | 'remover_espacos'
  | 'maiusculo'
  | 'minusculo'
  | 'formatar_cpf'
  | 'formatar_cnpj'
  | 'formatar_cep'
  | 'converter_data';

export interface ColumnMappingConfig {
  targetKey: string;
  sourceType: SourceType;
  sourceColumn: string; // Used if sourceType === 'column'
  fixedValue: string; // Used if sourceType === 'fixed'
  concatColumns: string[]; // Used if sourceType === 'concat'
  concatSeparator: string; // Used to join concatenated columns
  transformations: TransformationType[];
}

export const TARGET_COLUMNS: TargetColumnDefinition[] = [
  { key: 'id_da_venda', label: 'ID da Venda', description: 'Identificador único da transação', placeholder: 'Ex: 1001' },
  { key: 'data_do_pagamento', label: 'Data do Pagamento', description: 'Data da confirmação do pagamento', placeholder: 'Ex: 2026-05-30 ou 30/05/2026' },
  { key: 'id_do_produto', label: 'ID do Produto', description: 'Identificador único do produto vendido', placeholder: 'Ex: prod-99' },
  { key: 'nome_do_produto', label: 'Nome do Produto', description: 'Título comercial do item', placeholder: 'Ex: Curso de React Avançado' },
  { key: 'id_do_sku', label: 'ID do SKU', description: 'Código do SKU correspondente', placeholder: 'Ex: SKU-REACT-ADV' },
  { key: 'nome_do_sku', label: 'Nome do SKU', description: 'Variação ou nome específico do SKU', placeholder: 'Ex: Licença Individual' },
  { key: 'valor_unitario', label: 'Valor Unitário', description: 'Valor de um único item', placeholder: 'Ex: 299.90' },
  { key: 'moeda', label: 'Moeda', description: 'Símbolo monetário (geralmente BRL)', defaultFixed: 'BRL', placeholder: 'Ex: BRL' },
  { key: 'quantidade', label: 'Quantidade', description: 'Quantidade física ou licenças', placeholder: 'Ex: 1' },
  { key: 'nome_do_cliente', label: 'Nome do Cliente', description: 'Nome completo do comprador', placeholder: 'Ex: João Silva' },
  { key: 'email_do_cliente', label: 'E-mail do Cliente', description: 'Endereço de e-mail de contato', placeholder: 'Ex: joao@silva.com' },
  { key: 'telefone_do_cliente', label: 'Telefone do Cliente', description: 'Telefone com DDD', placeholder: 'Ex: 11999999999' },
  { key: 'tipo_de_cliente', label: 'Tipo de Cliente', description: 'Pessoa Física (PF) ou Jurídica (PJ)', defaultFixed: 'Pessoa Física', placeholder: 'Ex: Pessoa Física' },
  { key: 'documento_do_cliente', label: 'Documento do Cliente', description: 'CPF ou CNPJ do comprador', placeholder: 'Ex: 123.456.789-00' },
  { key: 'endereco_rua', label: 'Endereço (Rua)', description: 'Logradouro de entrega/faturamento', placeholder: 'Ex: Av. Paulista' },
  { key: 'endereco_numero', label: 'Endereço (Número)', description: 'Número residencial ou comercial', placeholder: 'Ex: 1000' },
  { key: 'endereco_bairro', label: 'Endereço (Bairro)', description: 'Bairro do endereço', placeholder: 'Ex: Bela Vista' },
  { key: 'endereco_complemento', label: 'Endereço (Complemento)', description: 'Sala, apto, bloco, etc.', placeholder: 'Ex: Apto 42' },
  { key: 'endereco_cidade', label: 'Endereço (Cidade)', description: 'Nome do município', placeholder: 'Ex: São Paulo' },
  { key: 'endereco_estado', label: 'Endereço (Estado)', description: 'Sigla do estado com 2 caracteres', placeholder: 'Ex: SP' },
  { key: 'endereco_cep', label: 'Endereço (CEP)', description: 'CEP formatado', placeholder: 'Ex: 01310-100' },
  { key: 'endereco_pais', label: 'Endereço (País)', description: 'País de origem do cliente', defaultFixed: 'Brasil', placeholder: 'Ex: Brasil' },
];

export interface ExcelRow {
  [key: string]: any;
}

export interface PreviewData {
  headers: string[];
  rows: ExcelRow[];
}
