package com.jb.erp.util;

import java.io.Serializable;
import java.util.Date;
import java.util.List;

import javax.enterprise.inject.Default;
import javax.inject.Inject;
import com.jb.erp.model.Despesa;
import com.jb.erp.repository.DespesasSearch;

@Default
public class ServiceDespesas implements Serializable{
	private static final long serialVersionUID = -1279340398511930041L;
	
	@Inject
	DespesasSearch despesasSearch;
	
	@Transacional
	public boolean salvarDespesa(Despesa Despesa) {
		return despesasSearch.save(Despesa);
	}
	
	@Transacional
	public List<Despesa> findDespesaWithUserId(Long userId){
		return despesasSearch.findDespesasWithUserId(userId);
	}

	@Transacional
	public boolean deleteLogico(Despesa despesa) {
		return despesasSearch.deletar(despesa);
	}
	
	@Transacional
	public List<Despesa> findWithTermo(String termo, Long userId) {
		return despesasSearch.findByDespesa(termo, userId);
	}
	
	@Transacional
	public List<Despesa> findWithTermo(String termo, Long userId, Date dataInit, Date dataFim) {
		return despesasSearch.findByDespesa(termo, userId, dataInit, dataFim);
	}
}
