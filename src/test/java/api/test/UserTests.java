package api.test;

import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.github.javafaker.Faker;

import api.endpoints.UserEndPoints;
import api.payloads.User;
import io.restassured.response.Response;

public class UserTests {
	Faker faker;
	User userPayload;
	
	@BeforeClass
	public void setupData() {
		faker = new Faker();
		userPayload = new User();
		userPayload.setId(faker.idNumber().hashCode());
		userPayload.setUsername(faker.name().username());
		userPayload.setFirstName(faker.name().firstName());
		userPayload.setLastName(faker.name().lastName());
		userPayload.setEmail(faker.internet().safeEmailAddress());
		userPayload.setPassword(faker.internet().password());
		userPayload.setPhone(faker.phoneNumber().cellPhone());
	}

	@Test(priority = 1)
	public void testPostUser() {
		Response response = UserEndPoints.createUser(userPayload);
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(), 200);
	}
	
	@Test(priority = 2)
	public void testUpdateByUser() {
		userPayload.setFirstName("Sujit");
		Response response = UserEndPoints.updateUser(userPayload.getUsername(),userPayload);
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(), 200);
		
		Response responseAfterUpdate = UserEndPoints.readUser(userPayload.getUsername());
		response.then().log().all();
		Assert.assertEquals(responseAfterUpdate.getStatusCode(), 200);
		
	}
	
	@Test(priority = 3)
	public void testGetUser() {
		Response response = UserEndPoints.readUser(userPayload.getUsername());
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(), 200);
	}
	
	@Test(priority = 4)
	public void testDeleteUser() {
		Response response = UserEndPoints.deleteUser(userPayload.getUsername());
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(), 200);
	}
}
