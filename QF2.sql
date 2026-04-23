SELECT     item, SUM(bal) AS quant, model
INTO            f_bal
FROM         bal_2006
GROUP BY item, model