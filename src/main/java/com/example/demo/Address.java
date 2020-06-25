package com.example.demo;

public class Address {

	private String street;
	private int pincode;

	public Address() {
	}

	public Address(String street, int pincode) {
		super();
		this.street = street;
		this.pincode = pincode;
	}

	public String getStreet() {
		return street;
	}

	public void setStreet(String street) {
		this.street = street;
	}

	public int getPincode() {
		return pincode;
	}

	public void setPincode(int pincode) {
		this.pincode = pincode;
	}

	@Override
	public String toString() {
		return "Address [street=" + street + ", pincode=" + pincode + "]";
	}

}
