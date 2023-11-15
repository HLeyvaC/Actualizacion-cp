CREATE OR REPLACE FUNCTION cat.new_cp()
RETURNS void AS 
$$
BEGIN

	UPDATE cat.nuevos
	SET d_tipo_asenta = cat.type_settlements.id::VARCHAR
	FROM cat.type_settlements
	WHERE UPPER(cat.nuevos.d_tipo_asenta) = cat.type_settlements.name;

	UPDATE cat.nuevos
	SET d_mnpio = cat.municipalities.id::VARCHAR
	FROM cat.municipalities
	WHERE cat.nuevos.d_mnpio = cat.municipalities.name;

	UPDATE cat.nuevos
	SET d_estado = cat.states.id::VARCHAR
	FROM cat.states
	WHERE cat.nuevos.d_estado = cat.states.name;

	UPDATE cat.nuevos
	SET d_ciudad = cat.localities.id::VARCHAR
	FROM cat.localities
	WHERE cat.nuevos.d_ciudad = cat.localities.name;


INSERT INTO cat.settlements (state_id, municipality_id, locality_id, type_settlement_id, postal_code, name, ambit) 
SELECT
    d_estado::INTEGER,
    d_mnpio::INTEGER,
    d_ciudad::INTEGER,
    d_tipo_asenta::INTEGER,
    d_codigo,
    d_asenta,
    d_zona
FROM cat.nuevos n
WHERE NOT EXISTS (
    SELECT 1
    FROM cat.settlements s
    WHERE s.postal_code = n.d_codigo
      AND s.name = n.d_asenta
);

END;
$$
LANGUAGE plpgsql;







