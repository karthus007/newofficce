package com.andin.thread;

import java.util.concurrent.Callable;

import com.andin.utils.OfficeFileUtil;

public class OfficeThread implements Callable<Boolean>{

	private String name;
	
	public OfficeThread(String name) {
		this.name = name;
	}
	
	@Override
	public Boolean call() throws Exception {
		return OfficeFileUtil.officeToPdf(name);
	}

}
