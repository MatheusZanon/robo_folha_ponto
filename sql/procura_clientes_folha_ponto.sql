SELECT *
FROM clientes_financeiro c
WHERE EXISTS (SELECT id FROM clientes_financeiro_folha_ponto cfp WHERE cfp.cliente_id = c.id) AND c.is_active = 1;