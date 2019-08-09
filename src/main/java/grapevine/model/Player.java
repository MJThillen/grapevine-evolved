package grapevine.model;

import grapevine.constants.Constants;

import java.io.ObjectInputStream;
import java.time.LocalDate;

public class Player {
    private int id;
    private String name;
    private String email;
    private String phone;
    private String position;
    private String status;
    private String address;
    private Experience playerExperience;
    private String notes;
    private LocalDate lastModified;

    public Player() {

    }

    public Player(int id, String name, String email, String phone, String position, String status, String address, Experience playerExperience, String notes) {
        this.id = id;
        this.name = name;
        this.email = email;
        this.phone = phone;
        this.position = position;
        this.status = status;
        this.address = address;
        this.playerExperience = playerExperience;
        this.notes = notes;
        this.lastModified = LocalDate.now();
    }

    public int outputId() {
        return Constants.OUTPUT_ID_CONSTANTS.none.getValue();
    }

    public void initializeForOutput() {
        Experience.initializeForOutput();
    }

    //ToDo: Input & Output functions

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public Experience getPlayerExperience() {
        return playerExperience;
    }

    public void setPlayerExperience(Experience playerExperience) {
        this.playerExperience = playerExperience;
    }

    public String getNotes() {
        return notes;
    }

    public void setNotes(String notes) {
        this.notes = notes;
    }

    public LocalDate getLastModified() {
        return lastModified;
    }

    public void setLastModified(LocalDate lastModified) {
        this.lastModified = lastModified;
    }

    public static Player inputFromBinary(ObjectInputStream inputStream, double version) {
        return null;
    }
}
