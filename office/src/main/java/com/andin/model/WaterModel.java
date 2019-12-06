package com.andin.model;

/**
 * PDF的水印信息
 * @author Administrator
 *
 */
public class WaterModel {

	/* --- 承办人 --- */
	private String handler;
	
	/* --- 部门负责人 --- */
	private String head;
	
	/* --- PDF查看密码 --- */
	private String pass;
	
	/* --- 公司名称 --- */
	private String com;
	
	/* --- 合同编号 --- */
	private String id;

	public String getHandler() {
		return handler;
	}

	public void setHandler(String handler) {
		this.handler = handler;
	}

	public String getHead() {
		return head;
	}

	public void setHead(String head) {
		this.head = head;
	}

	public String getPass() {
		return pass;
	}

	public void setPass(String pass) {
		this.pass = pass;
	}

	public String getCom() {
		return com;
	}

	public void setCom(String com) {
		this.com = com;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("WaterModel [handler=");
		builder.append(handler);
		builder.append(", head=");
		builder.append(head);
		builder.append(", pass=");
		builder.append(pass);
		builder.append(", com=");
		builder.append(com);
		builder.append(", id=");
		builder.append(id);
		builder.append("]");
		return builder.toString();
	}

	public WaterModel() {}
	
	public WaterModel(String handler, String head, String pass, String com, String id) {
		this.handler = handler;
		this.head = head;
		this.pass = pass;
		this.com = com;
		this.id = id;
	}
	
}
