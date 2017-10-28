package com.amazon.test;

import com.amazon.base.BaseClass;
import com.amazon.pages.DashboardPage;

public class DashboardPageTest extends BaseClass {
	DashboardPage dashboard;

	public DashboardPageTest() {
		super();
	}
	
	public void setUp() {
		initialization();
		dashboard = new DashboardPage();
	}
}
