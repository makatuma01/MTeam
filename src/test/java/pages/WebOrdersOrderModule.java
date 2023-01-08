package pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import utilities.Driver;

public class WebOrdersOrderModule {

    public WebOrdersOrderModule(){
        WebDriver driver = Driver.getDriver();
        PageFactory.initElements(driver,this);
    }

    @FindBy(id="ctl00_MainContent_fmwOrder_ddlProduct")
    public WebElement productDropdown;

    @FindBy(id="ctl00_MainContent_fmwOrder_txtQuantity")
    public WebElement quantityBox;

    @FindBy(id="ctl00_MainContent_fmwOrder_txtUnitPrice")
    public WebElement priceBox;

    @FindBy(id="ctl00_MainContent_fmwOrder_txtDiscount")
    public WebElement discountBox;

    @FindBy(id="ctl00_MainContent_fmwOrder_txtTotal")
    public WebElement totalBox;

}
