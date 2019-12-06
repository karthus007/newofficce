package com.andin.model;

public class TaskModel {

	/* --- 转换任务ID --- */
	private String id;
	
	/* --- 转换文件名称 --- */
	private String filename;
	
	/* --- 转换文件大小 --- */
	private Integer filesize;
	
	/* --- 转换文件类型：82:ppt类型 83:doc类型 84:pdf类型 85:xls类型  --- */
	private Integer filetype;
	
	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getFilename() {
		return filename;
	}

	public void setFilename(String filename) {
		this.filename = filename;
	}

	public Integer getFilesize() {
		return filesize;
	}

	public void setFilesize(Integer filesize) {
		this.filesize = filesize;
	}

	public Integer getFiletype() {
		return filetype;
	}

	public void setFiletype(Integer filetype) {
		this.filetype = filetype;
	}
	
	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("TaskModel [id=");
		builder.append(id);
		builder.append(", filename=");
		builder.append(filename);
		builder.append(", filesize=");
		builder.append(filesize);
		builder.append(", filetype=");
		builder.append(filetype);
		builder.append("]");
		return builder.toString();
	}

}
