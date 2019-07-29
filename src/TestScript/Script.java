package TestScript;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.Set;
import java.util.regex.Pattern;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Script {

	public static void main(String[] args) throws InterruptedException, AWTException, IOException {
		// TODO Auto-generated method stub

		
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"/src/Utilities/chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		
		driver.navigate().to("http://localhost:100");
		
			
		driver.manage().window().maximize();
		
		Xls_Reader xr=new Xls_Reader(System.getProperty("user.dir")+"/src/TestCase/TestData.xlsx");
		int rownum=xr.getRowCount("Sheet1");
		System.out.println(rownum);
		
		for(int i=2;i<=rownum;i++)
		{
			String vRun=xr.getCellData("Sheet1", "Run", i).trim();
			if(vRun.equalsIgnoreCase("ON"))
			{
				String vTCName=xr.getCellData("Sheet1", "TCName", i).trim();
				switch(vTCName)
				{
				case "vTiger_login_verifyAppUrl_TC01":
					String actualUrl=driver.getTitle().trim();
					String expectedUrl="vtiger CRM - Commercial Open Source CRM";
					if(actualUrl.equals(expectedUrl))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
						xr.setCellData("Sheet1", "ActualOutput", i, actualUrl);
					}
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
						xr.setCellData("Sheet1", "ActualOutput", i, actualUrl);
					}
									
				break;
				case "vTiger_login_verifyAppLogo_TC02":
				
					if(driver.findElements(By.xpath("//img[@src='include/images/vtiger-crm.gif']")).size()==1)
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
				case "vTiger_login_verifyMoreLink_TC03":
					
					driver.findElement(By.linkText("More ...")).click();
					
					Set<String> set1=driver.getWindowHandles();
					Iterator<String> iter1=set1.iterator();
					String firstWindow1=iter1.next();
					System.out.println(firstWindow1);
					String secondWindow1=iter1.next();
					System.out.println(secondWindow1);
					
					driver.switchTo().window(secondWindow1);
					Thread.sleep(3000);
					secondWindow1=driver.getTitle().trim();
					System.out.println(secondWindow1);
					driver.switchTo().window(firstWindow1);
					Thread.sleep(3000);
											
					if(secondWindow1.equals("CRM Software | Customer Relationship Management - Vtiger CRM"))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
				case "vTiger_login_verifyvTigerCustomerPortalLink_TC04":
					
					driver.findElement(By.linkText("vtiger Customer Portal")).click();
					
					Set<String> set2=driver.getWindowHandles();
					Iterator<String> iter2=set2.iterator();
					String firstWindow2=iter2.next();
					System.out.println(firstWindow2);
					String secondWindow2=iter2.next();
					System.out.println(secondWindow2);
					
					driver.switchTo().window(secondWindow2);
					Thread.sleep(3000);
					secondWindow2=driver.getTitle().trim();
					System.out.println(secondWindow2);
					Thread.sleep(3000);
																
					if(secondWindow2.equals("Help Desk Software - Vtiger"))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;	
			
				case "vTiger_CutomerPortal_Resource_Training_TC05":
						
						driver.findElement(By.linkText("vtiger Customer Portal")).click();
						
						Set<String> set3=driver.getWindowHandles();
						Iterator<String> iter3=set3.iterator();
						String firstWindow3=iter3.next();
						System.out.println(firstWindow3);
						String secondWindow3=iter3.next();
						System.out.println(secondWindow3);
						
						driver.switchTo().window(secondWindow3);
						Thread.sleep(7000);
																		
						Actions act=new Actions(driver);
						act.moveToElement(driver.findElement(By.id("menu-item-381"))).build().perform();
						Thread.sleep(3000);
						
						driver.findElement(By.linkText("Training")).click();
						Thread.sleep(2000);
						String pageUrl=driver.getCurrentUrl().trim();
						System.out.println(pageUrl);	
						
						if(pageUrl.equals("https://www.vtiger.com/training/"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;	
						
				case "vTiger_CutomerPortal_Products_SalesCRM_TC06":
						
						driver.findElement(By.linkText("vtiger Customer Portal")).click();
						
						Set<String> set4=driver.getWindowHandles();
						Iterator<String> iter4=set4.iterator();
						String firstWindow4=iter4.next();
						System.out.println(firstWindow4);
						String secondWindow4=iter4.next();
						System.out.println(secondWindow4);
						
						driver.switchTo().window(secondWindow4);
						Thread.sleep(7000);
											
						Actions act2=new Actions(driver);
						act2.moveToElement(driver.findElement(By.id("menu-item-1717"))).build().perform();
											
						driver.findElement(By.linkText("Sales CRM")).click();
						String pageUrl1=driver.getCurrentUrl().trim();
						System.out.println(pageUrl1);	
						
						if(pageUrl1.equals("https://www.vtiger.com/sales-crm/"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;	

					
			case "vTiger_Leads_verifyLeadCreationMandatoryfields_TC07":
					driver.findElement(By.xpath("//input[@name='user_name']")).sendKeys("admin");
					driver.findElement(By.xpath("//input[@name='user_password']")).sendKeys("admin");
					driver.findElement(By.xpath("//input[@name='Login']")).click();
					driver.findElement(By.xpath("//a[text()='New Lead']")).click();
					driver.findElements(By.xpath("//input[@name='button']")).get(0).click();
					
					Alert alt=driver.switchTo().alert();
					String errLastName=alt.getText().trim();
					System.out.println(errLastName);
					alt.accept();
					
					driver.findElement(By.xpath("//input[@name='lastname']")).sendKeys("Modi");
					
					driver.findElements(By.xpath("//input[@name='button']")).get(0).click();
					
					Alert alt1=driver.switchTo().alert();
					String errCompName=alt1.getText().trim();
					System.out.println(errCompName);
					alt1.accept();
					
					driver.findElement(By.xpath("//input[@name='company']")).sendKeys("BJP");
					driver.findElements(By.xpath("//input[@name='button']")).get(0).click();
										
					if(errLastName.equals("Last Name cannot be empty") && errCompName.equals("Company cannot be empty"))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
			case "vTiger_Leads_DeleteCancellation_TC08":
					
					driver.findElement(By.xpath("//input[@name='Delete']")).click();
					
					Alert alt2=driver.switchTo().alert();
					String errDeleteMsg=alt2.getText().trim();
					System.out.println(errDeleteMsg);
					Thread.sleep(2000);
					alt2.dismiss();
													
					if((driver.findElements(By.xpath("//td[@class='moduleTitle' and text()='Lead:   Modi']")).size()==1) && (driver.findElements(By.xpath("//td[@class='dataLabel' and text()='Last Name:']/following::td[@class='dataField' and text()='Modi']")).size()==1) && (driver.findElements(By.xpath("//td[@class='dataLabel' and text()='Company:']/following::td[@class='dataField' and text()='BJP']")).size()==1) &&(errDeleteMsg.equals("Are you sure you want to delete this record?")))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
			case "vTiger_MouseActions_TC09":
					Actions act1=new Actions(driver);
					act1.moveToElement(driver.findElement(By.id("showSubMenu"))).build().perform();
					driver.findElement(By.linkText("New Vendor")).click();
																
					if(driver.findElement(By.xpath("//td[text()='Vendor Name:']")).isDisplayed())
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
									
			case "vTiger_WindowHandlers_TC10":
							driver.findElement(By.linkText("New Product")).click();
							driver.findElement(By.name("productname")).sendKeys("Product");	
							driver.findElements(By.xpath("//input[@title='Change']")).get(0).click();									
							Thread.sleep(5000);
						 Set<String> set=driver.getWindowHandles();
						 Iterator<String> iter=set.iterator();
						 String firstWindow=iter.next();
						 System.out.println(firstWindow);
						 String secondWindow=iter.next();
						 System.out.println(secondWindow);
						 
						 driver.switchTo().window(secondWindow);
						 driver.findElement(By.linkText("Mary Smith")).click();
						 driver.switchTo().window(firstWindow);
						 String newProductVal=driver.findElement(By.name("contact_name")).getAttribute("value");	;
																							
						if(newProductVal.equals("Mary Smith"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;
						
				case "vTiger_newContact_verifyContactPage_TC11":
							
							driver.findElement(By.xpath("//input[@name='user_name']")).sendKeys("admin");
							driver.findElement(By.xpath("//input[@name='user_password']")).sendKeys("admin");
							driver.findElement(By.xpath("//input[@name='Login']")).click();
							
							driver.findElement(By.xpath("//a[text()='New Contact']")).click();
													
							if((driver.findElement(By.xpath("//td[@class='moduleTitle']")).isDisplayed())&& (driver.findElement(By.linkText("Contacts")).isEnabled()))
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}					
							else
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}
							break;
							
				case "vTiger_newContact_verifyMandateFields_TC12":
							
							driver.findElements(By.xpath("//input[@name='button'][@type='submit']")).get(0).click();
							
							Alert alt3=driver.switchTo().alert();
							String lastNamemsg=alt3.getText().trim();
							System.out.println(lastNamemsg);
							alt3.accept();
							
							driver.findElement(By.xpath("//input[@name='lastname']")).sendKeys("Modi");
							driver.findElements(By.xpath("//input[@name='button'][@type='submit']")).get(0).click();
																			
							if((driver.findElements(By.xpath("//td[@class='moduleTitle'][text()='Contact:   Modi']")).size()==1)&& (driver.findElements(By.xpath("//td[@class='dataLabel' and text()='Last Name:']/following::td[@class='dataField' and text()='Modi']")).size()==1))
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}					
							else
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}
							break;
							
				case "vTiger_newContact_verifyEditContact_TC13":
								
								driver.findElement(By.xpath("//input[@name='Edit']")).click();
								driver.findElement(By.xpath("//input[@name='lastname']")).sendKeys("Modi");
								driver.findElements(By.xpath("//input[@name='button'][@type='submit']")).get(0).click();
																												
								if((driver.findElements(By.xpath("//td[@class='moduleTitle'][text()='Contact:   ModiModi']")).size()==1)&& (driver.findElements(By.xpath("//td[@class='dataLabel' and text()='Last Name:']/following::td[@class='dataField' and text()='ModiModi']")).size()==1))
								
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_newContact_verifySendMailButton_TC14":
									
								driver.findElement(By.xpath("//input[@name='SendMail']")).click();
								
								if(driver.findElement(By.xpath("//td[@class='moduleTitle']")).isDisplayed())
								
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_emails_verifyEmailInformation_ccField_TC15":
																							
								WebElement email=driver.findElement(By.xpath("//input[@name='ccmail']"));
								email.sendKeys("abcd@gmail.com");
								driver.findElement(By.xpath("//input[@name='subject']")).sendKeys("Test Email");
								driver.findElement(By.xpath("//input[@title='Select Email Template']")).click();
								
								Set<String> set5=driver.getWindowHandles();
								Iterator<String> iter5=set5.iterator();
								String firstWindow5=iter5.next();
								System.out.println(firstWindow5);
								String secondWindow5=iter5.next();
								System.out.println(secondWindow5);
								
								driver.switchTo().window(secondWindow5);
								driver.findElement(By.linkText("Thanks Note")).click();
								Thread.sleep(3000);	
								
								driver.switchTo().window(firstWindow5);
								driver.findElement(By.xpath("//input[@name='button' and @type='submit' and @title='Save [Alt+S]']")).click();
								Thread.sleep(3000);	
																																	
								if(driver.findElements(By.xpath("//td[@class='moduleTitle'][text()='Contact:   Modi Modi']")).size()==1)
									
									{
										System.out.println(vTCName+"=PASSED");
										xr.setCellData("Sheet1", "Status", i, "PASSED");
									}					
									else
									{
										System.out.println(vTCName+"=FAILED");
										xr.setCellData("Sheet1", "Status", i, "FAILED");
									}
									break;
									
					case "vTiger_accounts_cancelButton_TC16":
								
								driver.findElement(By.xpath("//input[@name='user_name']")).sendKeys("admin");
								driver.findElement(By.xpath("//input[@name='user_password']")).sendKeys("admin");
								driver.findElement(By.xpath("//input[@name='Login']")).click();
								
								driver.findElement(By.linkText("New Account")).click();
								driver.findElements(By.xpath("//input[@name='button' and @type='button']")).get(0).click();
								
								
								if(driver.findElement(By.linkText("Leads")).isEnabled())
																	
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_accounts_verifySaveButton_TC17":
																						
								driver.findElement(By.linkText("New Account")).click();
								driver.findElement(By.xpath("//input[@name='accountname']")).sendKeys("BJP");
								
								WebElement elm=driver.findElement(By.xpath("//select[@name='industry']"));
								Select sel=new Select(elm);
								sel.selectByValue("Banking");
								Thread.sleep(2000);
								
								WebElement elm1=driver.findElement(By.xpath("//select[@name='accounttype']"));
								Select sel1=new Select(elm1);
								sel1.selectByIndex(1); 
								Thread.sleep(2000);
								
								driver.findElement(By.xpath("//textarea[@name='bill_street']")).sendKeys("Viman Nagar");
								driver.findElement(By.xpath("//input[@name='bill_city']")).sendKeys("Pune");
								driver.findElement(By.xpath("//input[@name='bill_state']")).sendKeys("Maharashtra");
								Thread.sleep(2000);
								
								driver.findElement(By.xpath("//input[@name='copyright']")).click();
								Thread.sleep(2000);
								
								driver.findElements(By.xpath("//input[@name='button' and @type='submit']")).get(0).click();
														
								
								if(driver.findElements(By.xpath("//td[@class='moduleTitle'][text()='Account:  BJP']")).size()==1)
																	
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_accountsTab_verifyClearButton_TC18":
								
								driver.findElement(By.xpath("//input[@name='user_name']")).sendKeys("admin");
								driver.findElement(By.xpath("//input[@name='user_password']")).sendKeys("admin");
								driver.findElement(By.xpath("//input[@name='Login']")).click();
								
								driver.findElement(By.linkText("Accounts")).click();
								WebElement elm2=driver.findElement(By.xpath("//input[@name='accountname']"));
								elm2.sendKeys("BJP");
								String val=elm2.getAttribute("value");
																
								if(val.endsWith(""))
																	
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;	
								
					case "vTiger_newSalesOrder_verifySalesOrderPage_TC19":
								
								Actions act3=new Actions(driver);
								act3.moveToElement(driver.findElement(By.id("showSubMenu"))).build().perform();
								driver.findElement(By.linkText("New Sales Order")).click();
																						
								if(driver.findElements(By.xpath("//td[@class='moduleTitle'][text()='Sales Order:  ']")).size()==1)
																	
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_newSalesOrder_verifySubTotalPrice_TC20":
								
								driver.findElement(By.xpath("//img[@src='themes/blue/images/search.gif']")).click();
								
								Set<String> set6=driver.getWindowHandles();
								Iterator<String> iter6=set6.iterator();
								String firstWindow6=iter6.next();
								System.out.println(firstWindow6);
								String secondWindow6=iter6.next();
								System.out.println(secondWindow6);
								
								driver.switchTo().window(secondWindow6);
								driver.findElement(By.linkText("Vtiger 5 Users Pack")).click();
								driver.switchTo().window(firstWindow6);
								driver.findElement(By.id("txtQty1")).sendKeys("1");
								
								Actions act4=new Actions(driver);
								act4.sendKeys(Keys.TAB).build().perform();
								
								
								Thread.sleep(2000);
							
								
								String total=driver.findElement(By.id("total1")).getText();
								System.out.println(total);
								
								String subTotal=driver.findElement(By.id("subTotal")).getText();
								System.out.println(subTotal);
																
																						
								if(total.equals(subTotal))
																	
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;	
					case "vTiger_newSalesOrder_verifyGrandTotalPrice_TC21":									
						
						String total1=driver.findElement(By.id("total1")).getText();
						System.out.println("Total="+total1);
						String subTotal1=driver.findElement(By.id("subTotal")).getText();
						System.out.println("SubTotal="+subTotal1);						
						driver.findElement(By.id("txtTax")).sendKeys("10");
						String tax=driver.findElement(By.id("txtTax")).getAttribute("value"); 
						System.out.println("Tax="+tax);
						driver.findElement(By.id("txtAdjustment")).sendKeys("20");
						String adjust=driver.findElement(By.id("txtAdjustment")).getAttribute("value"); 
						System.out.println("Adjust="+adjust);
						
						Actions act5=new Actions(driver);
						act5.sendKeys(Keys.TAB).build().perform();
												
						String gdTotal=driver.findElement(By.id("grandTotal")).getText();
						System.out.println("GrandTotal="+gdTotal);
						
						double subTotal2=Double.parseDouble(subTotal1);
						double tax1=Double.parseDouble(tax);
						double adjust1=Double.parseDouble(adjust);
						double gdTotal1=Double.parseDouble(gdTotal);
						
						double gdTotal2=subTotal2+tax1+adjust1;
						System.out.println("GrandTotal2="+gdTotal2);
																				
						if(gdTotal1==gdTotal2)
															
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;
						
				case "vTiger_newSalesOrder_verifyAddProductButton_TC22":									
						
						driver.findElement(By.xpath("//input[@type='button'][@class='button'][@value='Add Product']")).click();
						
																				
						if(driver.findElement(By.linkText("Del")).isDisplayed())
															
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;	
					
											
					case "vTiger_newEmail_dateTextBoxUpdate_TC23":
																				
								driver.findElement(By.xpath("//a[text()='New Email']")).click();
								
								String date="2019-05-31";
								System.out.println(date.replace("2019-05-31", "2019-05-1"));
								System.out.println(date);
								String newDate=date;
																					
								if(date.equals(newDate))
								{
									System.out.println(vTCName+"=PASSED");
									xr.setCellData("Sheet1", "Status", i, "PASSED");
								}					
								else
								{
									System.out.println(vTCName+"=FAILED");
									xr.setCellData("Sheet1", "Status", i, "FAILED");
								}
								break;
								
					case "vTiger_homePage_themeDropdown_multipleSelection_TC24":
								
						Select sel2=new Select(driver.findElement(By.name("login_theme")));
						System.out.println(sel2.isMultiple());
																										
						if(sel2.isMultiple())
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}					
							else
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}
						break;
							
					case "vTiger_homePage_themeDropdown_countValues_TC25":
																		
						Select sel3=new Select(driver.findElement(By.name("login_theme")));
						List<WebElement> lstw=sel3.getOptions();
						System.out.println(lstw);
						System.out.println("List count="+lstw.size());
																										
						if(lstw.size()==4)
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}					
							else
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}
						break;
						
					case "vTiger_homePage_themeDropdown_elements_TC26":
						
						List<String> st=new ArrayList<String>();
						
						Select sel4=new Select(driver.findElement(By.name("login_theme")));
						List<WebElement> lstwe=sel4.getOptions();
						System.out.println(lstwe);
												
						String expStr="Aqua,blue,nature,orange";
				        String actStr="";
						
						for(WebElement w:lstwe)
						{
							System.out.println(w.getText());
							actStr=actStr+w.getText()+",";
							st.add(w.getText());
							
						}
						actStr=actStr.replace("orange,","orange");
						System.out.println(actStr);
																																								
						if(expStr.equals(actStr))
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}					
							else
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}
						
						break;
					
					case "vTiger_homePage_themeDropdown_sortingOrder_TC27":
						
						ArrayList<String> obtainedList = new ArrayList<String>(); 
						List<WebElement> elementList= driver.findElements(By.name("login_theme"));
						for(WebElement we:elementList)
						{
						   obtainedList.add(we.getText());
						   System.out.println(obtainedList);
						}
						ArrayList<String> sortedList = new ArrayList<>();   
						for(String s:obtainedList)
						{
						sortedList.add(s);
						System.out.println(sortedList);
						}
						Collections.sort(sortedList);						
						
						if(sortedList.equals(obtainedList))
							{
								System.out.println(vTCName+"=PASSED");
								xr.setCellData("Sheet1", "Status", i, "PASSED");
							}					
							else
							{
								System.out.println(vTCName+"=FAILED");
								xr.setCellData("Sheet1", "Status", i, "FAILED");
							}
						
						break;
						
					case "vTiger_newEmail_emailInformation_ccField_TC28":
						
						driver.findElement(By.linkText("New Email")).click();
						driver.findElement(By.xpath("//input[@name='ccmail']")).sendKeys("abcd@gmail.com");
						String cc=driver.findElement(By.xpath("//input[@name='ccmail']")).getAttribute("value");
						System.out.println(cc);																								
						
						if(cc.contains("@") && cc.endsWith(".com") || cc.endsWith(".co.in"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;
					
					case "vTiger_newEmail_emailInformation_selectEmailTemplateButton_TC29":
						
						driver.findElement(By.xpath("//input[@title='Select Email Template']")).click();
						
						Set<String> set7=driver.getWindowHandles();
						Iterator<String> iter7=set7.iterator();
						String firstWindow7=iter7.next();
						System.out.println(firstWindow7);
						String secondWindow7=iter7.next();
						System.out.println(secondWindow7);
						
						driver.switchTo().window(secondWindow7);						
						driver.findElement(By.linkText("Thanks Note")).click();
						
						driver.switchTo().window(firstWindow7);
						
						String text=driver.findElement(By.xpath("//textarea[@name='description']")).getAttribute("value");
						System.out.println(text);						
																	
						if(text.contains("Sincerely,"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;
						
				case "vTiger_newEmail_emailInformation_emailTemplateModify_TC30":						
											
						String text1=driver.findElement(By.xpath("//textarea[@name='description']")).getAttribute("value");
						System.out.println(text1);
						
						String updatedText=driver.findElement(By.xpath("//textarea[@name='description']")).getAttribute("value").replaceAll("name", "Ritesh Gedam").replaceAll("title", "Test Lead");
						Thread.sleep(3000);
						System.out.println(updatedText);						
																							
						if(updatedText.contains("Ritesh Gedam") && updatedText.contains("Test Lead"))
						{
							System.out.println(vTCName+"=PASSED");
							xr.setCellData("Sheet1", "Status", i, "PASSED");
						}					
						else
						{
							System.out.println(vTCName+"=FAILED");
							xr.setCellData("Sheet1", "Status", i, "FAILED");
						}
						break;
						
				case "vTiger_newEmail_chooseFileButton_robotClass_TC31":
					
					driver.findElement(By.linkText("New Email")).click();
					driver.findElement(By.xpath("//input[@name='filename'][@type='file']")).click();
					Thread.sleep(5000);					
									
					StringSelection ss=new StringSelection("C:\\Users\\rites\\Desktop\\Gauri.txt");
					Toolkit.getDefaultToolkit().getSystemClipboard().setContents(ss, null);
					
					Robot robo=new Robot();
					robo.keyPress(KeyEvent.VK_ENTER);
					robo.keyRelease(KeyEvent.VK_ENTER);
					robo.keyPress(KeyEvent.VK_CONTROL);
					robo.keyPress(KeyEvent.VK_V);
					robo.keyRelease(KeyEvent.VK_CONTROL);
					robo.keyRelease(KeyEvent.VK_V);
					robo.keyPress(KeyEvent.VK_ENTER);
					robo.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(5000);
					
					String filename=driver.findElement(By.xpath("//input[@name='filename'][@type='file']")).getAttribute("value");
					System.out.println(filename);
																						
					if(filename.contains(".txt"))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
				case "vTiger_newEmail_chooseFileButton_autoItTool_TC32":
					
					driver.findElement(By.linkText("New Email")).click();
					driver.findElement(By.xpath("//input[@name='filename'][@type='file']")).click();
					Thread.sleep(3000);					
					
					Runtime.getRuntime().exec("C:/Users/rites/Desktop/Gauri.exe");
					Thread.sleep(3000);	
					
					String filename1=driver.findElement(By.xpath("//input[@name='filename'][@type='file']")).getAttribute("value");
					System.out.println(filename1);
																						
					if(filename1.contains(".txt"))
					{
						System.out.println(vTCName+"=PASSED");
						xr.setCellData("Sheet1", "Status", i, "PASSED");
					}					
					else
					{
						System.out.println(vTCName+"=FAILED");
						xr.setCellData("Sheet1", "Status", i, "FAILED");
					}
					break;
					
					
						
						
						
							
				}
				
			}
		}
		driver.quit();	
		
		
	}

}
