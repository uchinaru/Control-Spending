package com.jb.erp.controller;

import java.io.IOException;
import java.io.Serializable;
import java.util.Date;
import java.util.List;

import javax.annotation.PostConstruct;
import javax.enterprise.context.SessionScoped;
import javax.inject.Inject;
import javax.inject.Named;

import org.apache.poi.ss.usermodel.Workbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.jb.erp.model.Despesa;
import com.jb.erp.repository.TipoPagamentos;
import com.jb.erp.util.DateUtils;
import com.jb.erp.util.ExportUtils;
import com.jb.erp.util.ServiceDespesas;

@Named(value = "despesasListReport")
@SessionScoped
public class DespesasListReport extends FormAbstract implements Serializable{
	private static final long serialVersionUID = 1L;

	@Inject
	private DateUtils dateUtils;
	
	@Inject
	private ServiceDespesas serviceDespesas;
	
	@Inject
	private ExportUtils exportUtils;
	
	private Despesa despesa;
	private String nome;
	private String valor;
	private Date data;
	private String quantidade;
	private TipoPagamentos formaDePagamento;
	private List<Despesa> listarDespesas;
	private String descricao;
	private int total;
	
	private String termoPesquisa = "";
	private Date dataInicial = null;
	private Date dataFinal  = null;
	private String tituloArquivo = "Relatrio de despesas";
	private String[] tituloColunas = { "Nome", "Valor", "Data gasto", "Mes gasto", "Quantidade", "Forma de pagamento", "Descricao" };
	
	@PostConstruct
	public void init() {
		loadPageWhifUserLogged();
		listarDespesas();
	}
	
	public boolean validade() {
		
		if ("".equals(nome)) {
			getMessagesUtils().warning("Preencha o campo nome");
			return false;
		}

		if ("".equals(valor)) {
			getMessagesUtils().warning("Preencha o campo valor");
			return false;
		}

		if (data == null) {
			getMessagesUtils().warning("Preencha o campo data");
			return false;
		}

		if ("".equals(quantidade)) {
			getMessagesUtils().warning("Preencha o campo quantidade");
			return false;
		}
		
		if ("".equals(formaDePagamento)) {
			getMessagesUtils().warning("Preencha o campo Forma de pagamento");
			return false;
		}
		
		return true;
	}

	public void salvar() {
		
		if (validade()) {
			
			if(despesa != null && despesa.getId() != null){
				despesa.setId(despesa.getId());
			}else {
				despesa = new Despesa();
			}
			
			despesa.setNome(nome);
			despesa.setValor(Double.parseDouble(valor.replace(",", ".")));
			despesa.setDataCusto(dateUtils.transformaDataSimples(data));
			despesa.setMesGasto(dateUtils.getMes(data));
			despesa.setQuantidade(Integer.parseInt(quantidade));
			despesa.setTipoPagamentos(formaDePagamento);
			despesa.setDescricao(descricao);
			despesa.setDeletado(false);
			despesa.setUserId(userLogged());
			
			if(serviceDespesas.salvarDespesa(despesa)) {
				getMessagesUtils().info("Registro salvo com sucesso");
			}else {
				getMessagesUtils().error("Erro ao salvar");
			}
			clearDados();
		}
	}
	
	public void listarDespesas() {
		
		listarDespesas = serviceDespesas.findDespesaWithUserId(userLogged());
		
		if(listarDespesas.isEmpty())
			getMessagesUtils().warning("Sem dados registrados");
			
		total = listarDespesas.size();
	}
	
	public void clearDados() {
		nome = null;
		valor = null;
		data = null;
		quantidade = null;
		formaDePagamento = null;
		descricao = null;
		dataInicial = null;
		dataFinal = null;
		despesa = null;
	}
	
	public boolean isDespesaSelected() {
		return despesa != null && despesa.getId() != null;
	}

	public void despesaSelecionada() {

		if (despesa != null && despesa.getId() != null) {
			
			nome = despesa.getNome();
			valor = String.valueOf(despesa.getValor());
			data = despesa.getDataCusto();
			quantidade = String.valueOf(despesa.getQuantidade());
			formaDePagamento = despesa.getTipoPagamentos();
			descricao = despesa.getDescricao();		
			
		} else {
			getMessagesUtils().warning("Selecione uma despesa");
		}
	}
	
	public void deletarDespesa() {
		
		if (serviceDespesas.deleteLogico(despesa)) {
			getMessagesUtils().info("Registro deletado!");
		}else {
			getMessagesUtils().warning("Algo deu errado.");
		}
		
	}
	
	public void pesquisar() {
		
		if(listarDespesas != null)
		listarDespesas.clear();
		
		if(dataInicial != null && dataFinal != null) {
			listarDespesas = serviceDespesas.findWithTermo(termoPesquisa, userLogged(), dataInicial, dataFinal);
		} else {
			listarDespesas = serviceDespesas.findWithTermo(termoPesquisa, userLogged());
		}
		
		total = listarDespesas.size();
		termoPesquisa = "";
	}
	
	
	public Workbook exportExcel() throws IOException {
		
		if (listarDespesas != null) {
			
			return exportUtils.exportExcelDespesa(tituloArquivo, tituloColunas, listarDespesas);
			
		} else {
			getMessagesUtils().warning("É necessário efetuar uma pesquisa");
			return null;
		}

	}

	public Workbook exportCsv() throws IOException {
		
		if (listarDespesas != null) {
			
			return exportUtils.exportCSVDespesa(tituloArquivo, tituloColunas, listarDespesas);
			
		} else {
			getMessagesUtils().warning("É necessário efetuar uma pesquisa");
			return null;
		}

	}
	
	public void exportPDF() throws IOException, DocumentException {
		
		if (listarDespesas != null) {

			exportUtils.exportPDFDespesa(tituloArquivo, tituloColunas, listarDespesas);

		} else {
			getMessagesUtils().warning("É necessário efetuar uma pesquisa");
		}

	}
	
	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	public String getValor() {
		return valor;
	}

	public void setValor(String valor) {
		this.valor = valor;
	}

	public Date getData() {
		return data;
	}

	public void setData(Date data) {
		this.data = data;
	}

	public String getQuantidade() {
		return quantidade;
	}

	public void setQuantidade(String quantidade) {
		this.quantidade = quantidade;
	}

	public String getDescricao() {
		return descricao;
	}

	public void setDescricao(String descricao) {
		this.descricao = descricao;
	}

	public Despesa getDespesa() {
		return despesa;
	}

	public void setDespesa(Despesa despesa) {
		this.despesa = despesa;
	}
	
	public TipoPagamentos[] getTipoPagamentos() {
		return TipoPagamentos.values();
	}

	public TipoPagamentos getFormaDePagamento() {
		return formaDePagamento;
	}

	public void setFormaDePagamento(TipoPagamentos formaDePagamento) {
		this.formaDePagamento = formaDePagamento;
	}

	public int getTotal() {
		return total;
	}

	public void setTotal(int total) {
		this.total = total;
	}

	public List<Despesa> getListarDespesas() {
		return listarDespesas;
	}

	public void setListarDespesas(List<Despesa> listarDespesas) {
		this.listarDespesas = listarDespesas;
	}

	public String getTermoPesquisa() {
		return termoPesquisa;
	}

	public void setTermoPesquisa(String termoPesquisa) {
		this.termoPesquisa = termoPesquisa;
	}

	public Date getDataInicial() {
		return dataInicial;
	}

	public void setDataInicial(Date dataInicial) {
		this.dataInicial = dataInicial;
	}

	public Date getDataFinal() {
		return dataFinal;
	}

	public void setDataFinal(Date dataFinal) {
		this.dataFinal = dataFinal;
	}

}
