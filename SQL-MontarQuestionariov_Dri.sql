select T2.Valor, QV.Id 'Quest Versao ID', QV.Identificador, Q2.DataCriacao, QuestionarioId, T.Valor, T.Codigo, QV.* 
from Questionario2 Q2
	inner join QuestionarioVersao QV
		on QV.QuestionarioId = Q2.Id
	inner join Termo T2
		on Q2.Codigo = T2.Codigo and T2.IdiomaId =1
	inner join Termo T
		on T.Codigo = QV.Codigo	and T.IdiomaId = 1
where
QV.DataCriacao > '2020-04-01'
--qv.Id = '546FF7E7-F82D-43D1-AF14-4192D1CB62FC'
--and T2.Valor like '%Requisito%'
order by T.Valor, QV.DataCriacao desc


--- ## MONTAR QUESTIONÁRIO
DECLARE @qvid as uniqueidentifier = '026CB61F-E290-4F59-8CEB-93D43BF9A9E5'


insert RelQuestionarioPergunta (id, PerguntaVersaoId, QuestionarioVersaoId, UsuarioId, DataCriacao, UsuarioIdUltimaAlteracao, DataUltimaAlteracao, 
ordem, Obrigatoria, peso, Pex)
--select top 1 * from RelQuestionarioPergunta order by DataCriacao desc

select NEWID(), PV.Id, @qvid,'4FC9956E-F874-4155-AA58-51B18EDE1F3D', GETDATE(), NULL, NULL, 
RANK() over (order by CAST(p2.Identificador AS VARCHAR)), 0 , 1.0, 0
from PerguntaVersao PV
	inner join Pergunta2 P2
		on P2.id = PV.PerguntaId
	inner join CategoriaPergunta CP
		on P2.CategoriaPerguntaId = CP.Id
	inner join Termo T
		on T.Codigo = CP.Codigo and T.IdiomaId = 1
where
P2.Identificador between 11220 and 11238	
order by p2.Identificador asc


11220-11238




select * from PerguntaVersao
where Id = 'C8C50BE7-8F17-40B1-AED0-D490823EC189'


delete from RelQuestionarioPergunta
where Id in (
select RQV.ID from RelQuestionarioPergunta RQV
	inner join PerguntaVersao PV
		ON RQV.PerguntaVersaoId = PV.Id
	inner join Pergunta2 P
		on PV.PerguntaId = p.Id
where p.Identificador between 11222 and 11238)

11220 and 11238
