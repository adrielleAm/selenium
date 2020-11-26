SELECT 
T.Valor 'TipoAuditoria', CQ.CategoriaQuestionario,Q.nome 'Q', Qv.nome 'Questionario', CA.Ano, CA.NumeroCiclo, CA.Nome 'Mês', DataInicio, DataFim 
 FROM CicloAuditoria CA
	inner join CategoriaQuestionario Cq
		on CQ.IDCategoriaQuestionario = CA.IDTipoAuditoria
	inner join TipoLoteAuditoria TLA
		on CQ.TipoLoteAuditoriaId = TLA.Id
	inner join Termo T
		on T.Codigo = TLA.codigo and T.IdiomaId = 1
	inner join RelCicloQuestionario RCQ
		on RCQ.CicloAuditoriaId = CA.IDCicloAuditoria
	inner join QuestionarioVersao Qv
		on Qv.Id = RCQ.QuestionarioVersaoId
	inner join Questionario2 q
		on Qv.QuestionarioId = Q.Id
where CA.DataCriacao > '2020-05-01' 
order by T.Valor, CQ.CategoriaQuestionario, Qv.nome, NumeroCiclo

select T.Valor, Q.Id, Q.Nome 'q', Qv.id, Qv.Nome 'QV'  from  QuestionarioVersao QV
	inner join Questionario2 q
		on Qv.QuestionarioId = Q.Id
	inner join Termo T
		on T.Codigo = Q.codigo and T.IdiomaId = 1
where qv.DataCriacao > '2020-04-27'
and (Q.Nome = 'PILAR FINANCEIRO PY SPO 2020' or Q.nome = 'PILAR SEGURANÇA BOL SPO 2020')

order by Q.Nome

UPDATE QuestionarioVersao
SET NOME = 'PILAR SEGURANÇA BOL SPO 2020'
WHERE id = '3955B17B-5ED9-40C2-A3B7-F095E0645B04'

UPDATE Questionario2
SET NOME = 'PILAR SEGURANÇA BOL SPO 2020'
WHERE id = 'BD1A3985-FC00-4052-BA29-1369F9F066BF'
