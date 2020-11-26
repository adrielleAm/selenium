select DISTINCT P.NomePerfil, CQ.CategoriaQuestionario from TBPerfil P
	inner join TBPerfilCategoriaQuestionario PCQ
		on PCQ.IDPerfil = P.IDPerfil
	inner join CategoriaQuestionario CQ
		on CQ.IDCategoriaQuestionario = PCQ.IDCategoriaQuestionario
	where CategoriaQuestionario = '1 - GENTE AS'
	and pCQ.Ativo = 1
	and Cq.TipoLoteAuditoriaId in (
	'FC4B1039-4238-495D-B85A-582237E6E6B1',
	'85D7954E-6261-4AEB-BABD-05E010009456',
	'691D8416-7C95-40C3-AB20-249E86A4B9BE')


	select TL.ID,T.Valor from TipoLoteAuditoria TL
		inner join Termo T 
			on T.Codigo = TL.Codigo and T.IdiomaId = 1
			order by T.Valor
