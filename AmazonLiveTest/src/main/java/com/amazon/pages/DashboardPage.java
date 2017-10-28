package com.amazon.pages;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.amazon.base.BaseClass;

public class DashboardPage extends BaseClass {
	// Initializing the Page Objects:
	public DashboardPage() {
		PageFactory.initElements(driver, this);
	}

	/* Variables */

	/* Locators */
	@FindBy(id = "nav-link-yourAccount")
	private WebElement signIn;

	/* Methods */
	public LoginPage clickSignIn() {
		signIn.click();
		return new LoginPage();
	}
}
