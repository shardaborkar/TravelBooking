package airportTDD.flightBookingTDD.tests;

import java.io.IOException;

import org.testng.Assert;
import org.testng.ITestListener;
import org.testng.ITestResult;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import airportTDD.flightBookingTDD.testcomponent.BaseTest;
import airportTDD.pageComponents.MultiTrip;
import airportTDD.pageComponents.OneWAY;
import airportTDD.pageComponents.RoundTrip;
import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import mainPageObject.TravelHomePage;

public class TopSectionTest extends BaseTest {

	@Test()
	public void searchWithMulticity() throws InterruptedException, IOException {
		
		System.out.println(travelHomePage.topSection().getFlightAttribute());
	}}