package grapevine.service;

import com.sun.istack.NotNull;
import grapevine.constants.Constants;
import grapevine.constants.ListDisplay;
import grapevine.constants.RumorCategory;
import grapevine.model.*;
import grapevine.model.Character;

import java.io.ObjectInputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedList;

import static grapevine.constants.QueryConstants.*;

public class APREngine {
    private Game game;
    private LinkedList<Action> actions;
    private LinkedList<Plot> plots;
    private LinkedList<Rumor> rumors;
    private TraitList<Trait> bgActions;
    private TraitList<Trait> apLevel;

    private boolean addCommon;
    private boolean carryUnused;
    private boolean publicRumors;
    private boolean personalRumors;
    private boolean raceRumors;
    private boolean groupRumors;
    private boolean subGroupRumors;
    private boolean influenceRumors;
    private boolean previousRumors;
    private boolean copyPrevious;

    private int personalActions;


    public APREngine(@NotNull final Game game) {
        this.game = game;
        actions = new LinkedList<>();
        plots = new LinkedList<>();
        rumors = new LinkedList<>();
        bgActions = new TraitList<Trait>("Backgrounds", false, false, false, ListDisplay.SIMPLE);
        apLevel = new TraitList<Trait>("Actions", false, false, false, ListDisplay.SIMPLE);
        initialize();
    }

    public static APREngine inputFromBinary(ObjectInputStream inputStream) throws ClassNotFoundException {
        return null;
    }

    public LinkedList<Action> getActions() {
        return actions;
    }

    public void setActions(LinkedList<Action> actions) {
        this.actions = actions;
    }

    public LinkedList<Plot> getPlots() {
        return plots;
    }

    public void setPlots(LinkedList<Plot> plots) {
        this.plots = plots;
    }

    public LinkedList<Rumor> getRumors() {
        return rumors;
    }

    public void setRumors(LinkedList<Rumor> rumors) {
        this.rumors = rumors;
    }

    public TraitList<Trait> getBgActions() {
        return bgActions;
    }

    public void setBgActions(TraitList<Trait> bgActions) {
        this.bgActions = bgActions;
    }

    public TraitList<Trait> getApLevel() {
        return apLevel;
    }

    public void setApLevel(TraitList<Trait> apLevel) {
        this.apLevel = apLevel;
    }

    public boolean isAddCommon() {
        return addCommon;
    }

    public void setAddCommon(boolean addCommon) {
        this.addCommon = addCommon;
    }

    public boolean isCarryUnused() {
        return carryUnused;
    }

    public void setCarryUnused(boolean carryUnused) {
        this.carryUnused = carryUnused;
    }

    public boolean isPublicRumors() {
        return publicRumors;
    }

    public void setPublicRumors(boolean publicRumors) {
        this.publicRumors = publicRumors;
    }

    public boolean isPersonalRumors() {
        return personalRumors;
    }

    public void setPersonalRumors(boolean personalRumors) {
        this.personalRumors = personalRumors;
    }

    public boolean isRaceRumors() {
        return raceRumors;
    }

    public void setRaceRumors(boolean raceRumors) {
        this.raceRumors = raceRumors;
    }

    public boolean isGroupRumors() {
        return groupRumors;
    }

    public void setGroupRumors(boolean groupRumors) {
        this.groupRumors = groupRumors;
    }

    public boolean isSubGroupRumors() {
        return subGroupRumors;
    }

    public void setSubGroupRumors(boolean subGroupRumors) {
        this.subGroupRumors = subGroupRumors;
    }

    public boolean isInfluenceRumors() {
        return influenceRumors;
    }

    public void setInfluenceRumors(boolean influenceRumors) {
        this.influenceRumors = influenceRumors;
    }

    public boolean isPreviousRumors() {
        return previousRumors;
    }

    public void setPreviousRumors(boolean previousRumors) {
        this.previousRumors = previousRumors;
    }

    public boolean isCopyPrevious() {
        return copyPrevious;
    }

    public void setCopyPrevious(boolean copyPrevious) {
        this.copyPrevious = copyPrevious;
    }

    public int getPersonalActions() {
        return personalActions;
    }

    public void setPersonalActions(int personalActions) {
        this.personalActions = personalActions;
    }

    public void initialize() {
        actions.clear();
        plots.clear();
        rumors.clear();
        bgActions.clear();
        apLevel.clear();

        bgActions.add(new Trait("Contacts"));
        bgActions.add(new Trait("Resources"));
        for(int i = 1; i <= 10; i++) {
            apLevel.add(new Trait("" + i, 2*i, ""));
        }

        addCommon = false;
        carryUnused = false;
        publicRumors = true;
        personalRumors = false;
        raceRumors = false;
        groupRumors = false;
        subGroupRumors = false;
        influenceRumors = true;
        previousRumors = true;
        copyPrevious = false;
    }

    public void addStandardRumors(final LocalDate when) {
        ArrayList<String> existing = new ArrayList<>();
        Rumor newRumor;
        String newName;

        //ToDo: In the original code, it turns the pointer to an hourglass here.

        for(Rumor rumor : rumors) {
            if (rumor.getDate() != null && rumor.getDate().isEqual(when)) {
                existing.add(rumor.getName());
            }
        }
        if (previousRumors) { //Add previously existing rumors to this date
            for(Rumor rumor : rumors) {
                if (rumor.getDate() != null && rumor.getDate().isBefore(when)) {
                    newRumor = (Rumor) rumor;
                    newRumor.setDate(when);
                    rumors.add(newRumor);
                    existing.add(newRumor.getName());
                }
            }
        }
        if (publicRumors) {
            if (!existing.contains(Constants.PUBLIC_RUMOR_TITLE)) {
                newRumor = new Rumor();
                newRumor.setName(Constants.PUBLIC_RUMOR_TITLE);
                newRumor.setDate(when);
                newRumor.setCategory(RumorCategory.GENERAL);
                rumors.add(newRumor);
                existing.add(newRumor.getName());
            }
        }

        for(Character character : game.getCharacters()) {
            if (character.isActive()) {
                newName = character.getName();
                if (personalRumors && !existing.contains(newName)) {
                    newRumor = new Rumor();
                    newRumor.initializeQuery(newName, when, RumorCategory.PERSONAL);
                    newRumor.addClauseToQuery(QueryKeys.NAME,
                            newName,
                            0,
                            QueryCompare.EQUALS,
                            false);
                    rumors.add(newRumor);
                    existing.add(newName);
                }
                newName = character.getRace();
                if (raceRumors && !existing.contains(newName)) {
                    newRumor = new Rumor();
                    newRumor.initializeQuery(newName, when, RumorCategory.RACE);
                    newRumor.addClauseToQuery(QueryKeys.RACE,
                            newName,
                            0,
                            QueryCompare.EQUALS,
                            false);
                    rumors.add(newRumor);
                    existing.add(newName);
                }
                newName = character.getGroup();
                if (groupRumors && !existing.contains(newName)) {
                    newRumor = new Rumor();
                    newRumor.initializeQuery(newName, when, RumorCategory.GROUP);
                    newRumor.addClauseToQuery(QueryKeys.GROUP,
                            newName,
                            0,
                            QueryCompare.EQUALS,
                            false);
                    rumors.add(newRumor);
                    existing.add(newName);
                }
                newName = character.getSubGroup();
                if (subGroupRumors && !existing.contains(newName)) {
                    newRumor = new Rumor();
                    newRumor.initializeQuery(newName, when, RumorCategory.SUBGROUP);
                    newRumor.addClauseToQuery(QueryKeys.SUBGROUP,
                            newName,
                            0,
                            QueryCompare.EQUALS,
                            false);
                    rumors.add(newRumor);
                    existing.add(newName);
                }

                if (influenceRumors) {
                    TraitList<Trait> influences = character.getValue(QueryKeys.INFLUENCES);
                    for (Trait influence : influences) {
                        newName = influence.getName() + " Influence";
                        if (!existing.contains(newName)) {
                            newRumor = new Rumor();
                            newRumor.influenceRumorSetup(
                                    newName, when, QueryKeys.INFLUENCES.getValue(), influence.getName());
                            rumors.add(newRumor);
                            existing.add(newName);
                        }
                    }
                }
            }
        }
        //ToDo: In the original code, the cursor is set back to normal here.
    }


}
