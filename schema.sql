-- Sistema Financeiro em Python e Excel Heitor Jorge
-- Estrutura relacional e dados iniciais de demonstração.

CREATE DATABASE IF NOT EXISTS sistema_financeiro
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE sistema_financeiro;

CREATE TABLE IF NOT EXISTS produtos (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    categoria VARCHAR(60) NOT NULL,
    preco_unitario DECIMAL(10, 2) NOT NULL,
    estoque_inicial INT NOT NULL DEFAULT 0,
    criado_em TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CHECK (preco_unitario >= 0),
    CHECK (estoque_inicial >= 0)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS movimentacoes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    produto_id INT NOT NULL,
    tipo ENUM('Entrada', 'Saida') NOT NULL,
    quantidade INT NOT NULL,
    data_movimento DATE NOT NULL,
    valor_unitario DECIMAL(10, 2) NOT NULL,
    observacao VARCHAR(255) NOT NULL DEFAULT '',
    CONSTRAINT fk_movimentacoes_produtos
        FOREIGN KEY (produto_id) REFERENCES produtos(id)
        ON UPDATE CASCADE
        ON DELETE RESTRICT,
    CHECK (quantidade > 0),
    CHECK (valor_unitario >= 0),
    INDEX idx_movimentacoes_data (data_movimento),
    INDEX idx_movimentacoes_tipo (tipo)
) ENGINE=InnoDB;

-- Produtos iniciais (só são inseridos se a tabela estiver vazia)
INSERT INTO produtos (nome, categoria, preco_unitario, estoque_inicial)
SELECT d.nome, d.categoria, d.preco_unitario, d.estoque_inicial
FROM (
    SELECT 'Teclado mecanico' AS nome, 'Perifericos' AS categoria, 189.90 AS preco_unitario, 15 AS estoque_inicial
    UNION ALL SELECT 'Mouse sem fio', 'Perifericos', 79.90, 30
    UNION ALL SELECT 'Monitor 24 polegadas', 'Monitores', 749.00, 6
    UNION ALL SELECT 'Cabo HDMI 2m', 'Acessorios', 22.50, 60
    UNION ALL SELECT 'Hub USB 4 portas', 'Acessorios', 64.90, 18
) AS d
WHERE NOT EXISTS (SELECT 1 FROM produtos);

-- Movimentações iniciais (só são inseridas se a tabela estiver vazia)
INSERT INTO movimentacoes
    (produto_id, tipo, quantidade, data_movimento, valor_unitario, observacao)
SELECT p.id, d.tipo, d.quantidade, d.data_movimento, d.valor_unitario, d.observacao
FROM (
    SELECT 'Teclado mecanico' AS produto, 'Entrada' AS tipo, 10 AS quantidade, '2026-02-03' AS data_movimento, 189.90 AS valor_unitario, 'Compra de fornecedor' AS observacao
    UNION ALL SELECT 'Mouse sem fio', 'Saida', 12, '2026-02-10', 79.90, 'Venda'
    UNION ALL SELECT 'Monitor 24 polegadas', 'Entrada', 4, '2026-02-24', 749.00, 'Reposicao de estoque'
    UNION ALL SELECT 'Cabo HDMI 2m', 'Saida', 25, '2026-03-05', 22.50, 'Venda'
    UNION ALL SELECT 'Hub USB 4 portas', 'Entrada', 12, '2026-03-12', 64.90, 'Compra de fornecedor'
    UNION ALL SELECT 'Teclado mecanico', 'Saida', 8, '2026-03-27', 189.90, 'Venda'
    UNION ALL SELECT 'Mouse sem fio', 'Entrada', 20, '2026-04-06', 79.90, 'Reposicao de estoque'
    UNION ALL SELECT 'Monitor 24 polegadas', 'Saida', 3, '2026-04-18', 749.00, 'Venda'
    UNION ALL SELECT 'Cabo HDMI 2m', 'Entrada', 40, '2026-05-04', 22.50, 'Compra de fornecedor'
    UNION ALL SELECT 'Hub USB 4 portas', 'Saida', 9, '2026-05-16', 64.90, 'Venda'
    UNION ALL SELECT 'Teclado mecanico', 'Entrada', 6, '2026-06-08', 189.90, 'Reposicao de estoque'
    UNION ALL SELECT 'Mouse sem fio', 'Saida', 18, '2026-06-22', 79.90, 'Venda'
    UNION ALL SELECT 'Monitor 24 polegadas', 'Entrada', 5, '2026-07-09', 749.00, 'Compra de fornecedor'
    UNION ALL SELECT 'Cabo HDMI 2m', 'Saida', 30, '2026-07-21', 22.50, 'Venda'
    UNION ALL SELECT 'Hub USB 4 portas', 'Entrada', 8, '2026-08-13', 64.90, 'Reposicao de estoque'
) AS d
JOIN produtos AS p ON p.nome = d.produto
WHERE NOT EXISTS (SELECT 1 FROM movimentacoes);
