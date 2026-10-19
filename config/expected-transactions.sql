SELECT
    transaction_datetime AS `Transaction Date/Time`,
    merchant_reference AS `Merchant Reference`,
    payment_auth_id AS `Payment / Auth ID`,
    channel AS `Channel`,
    card_scheme AS `Card Scheme`,
    masked_pan AS `Masked PAN`,
    transaction_type AS `Transaction Type`,
    status AS `Status`,
    status2 AS `status2`,
    gross_amount AS `Gross Amount`,
    refund AS `Refund`,
    processing_fee AS `Processing Fee`,
    net_amount AS `Net Amount`,
    settlement_batch AS `Settlement Batch`,
    '' AS `TESTIJNG`,
    '' AS `sfasdfds`,
    '' AS `sadfwe3`
FROM bulk_validator.transactions;
