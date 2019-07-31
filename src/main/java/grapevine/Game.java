package grapevine;

import grapevine.constants.ExperienceChange;
import grapevine.constants.FileFormat;
import grapevine.constants.Race;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import grapevine.constants.Constants;
import grapevine.util.ImportHelper;

public class Game {
    private List<Player> players; //Collection of all players
    private List<Character> characters; //collection of all characters
    private List<Item> items; //collection of all items
    private List<Rote> rotes; //collection of all rotes
    private List<Location> locations; //collection of all locations
    private List<List<Rumor>> rumorLists; //collection of lists of rumors, one per date
    private List<InfluenceAction> influenceUses; //collection of influence use actions
    private MenuSet menuSet; //collection of menus
    private QueryEngine queryEngine; //collection, logic of queries
    private APREngine aprEngine; //collection, logic of actions/plots/rumors

    private FileFormat fileFormat; //current file type to save file as
    private String gameFilename; //full pathname of game file

    private String chronicleTitle; //Title of the Chronicle
    private String website; //grapevine.Game URL
    private String email; //Main ST Email address
    private String phone; //Main Phone Number
    private String usualSite; //Usual grapevine.Game Site
    private String usualTime; //Usual grapevine.Game Start Time
    private GameCalendar gameCalendar; //Calendar of game dates
    private String description; //Description of your game
    private boolean extendedHealth; //whether this game uses abbreviated or extended health levels
    private boolean enforceHistory; //whether to enforce use of XP history
    private boolean linkTraitMaxes; //Whether to link trait maximums on character sheets
    private String randomTraits; //Comma-separated list of random trait options
    private String stCommentStart; //opening markup of an ST comment
    private String stCommentEnd; //closing markup of an ST comment
    private List<ExperienceAward> xpAwardList; //list of standard XP and PP awards
    private List<Template> outputTemplates; //list of output templates
    private boolean fileError; //whether a file error happened during open or save
    private String errorMessage; //description of the error
    private String mergeResults; //line-delimited results of a merge or exchange file load
    private ProgressBar fileProgress; //control describing progress of load
    private DuplicateAction duplicateAction; //What action to take when duplicating characters
    private boolean duplicateAll; //Whether to take that action in all cases

    //Percent of progress bar to fill for each loading part
    private static final int CALENDAR_PERCENT = 5;
    private static final int PLAYER_PERCENT = 30;
    private static final int CHARACTER_PERCENT = 65;

    private boolean changed; //Dirty Flag

    /**
     * Create and initialize all needed objects.
     */
    public Game() {
        this.players = new ArrayList<>();
        this.characters = new ArrayList<>();
        this.items = new ArrayList<>();
        this.rotes = new ArrayList<>();
        this.locations = new ArrayList<>();
        this.rumorLists = new ArrayList<>();
        this.influenceUses = new ArrayList<>();
        this.menuSet = new MenuSet();
        this.queryEngine = new QueryEngine();
        this.aprEngine = new APREngine();
        this.gameCalendar = new GameCalendar();
        this.xpAwardList = new ArrayList<>();
        this.outputTemplates = new ArrayList<>();

        extendedHealth = true;
        enforceHistory = true;
        linkTraitMaxes = true;

        randomTraits = "7,5,3,5,5,5,5";
        stCommentStart = "[ST]";
        stCommentEnd = "[/ST]";
    }

    /**
     * Adds the default output templates to this game.
     */
    private void addDefaultTemplates() {
        Template placeholder = new Template();
        for (Race race : Race.values()) {
            placeholder.setCharacterSheet(true);
            placeholder.setName(race.getName() + Constants.TN_CHARACTER_SHEET_SUFFIX);
            outputTemplates.add(placeholder);
        }
        //ToDo: Actually improve this process.
        placeholder.setCharacterSheet(false);
        placeholder.setName(Constants.TN_ACTION_RUMOR);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_MASTER_ACTION);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_MASTER_RUMOR);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_PLOT);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_CHARACTER_SHEETS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_CHARACTER_ROSTER);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_EQUIPMENT);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_SIGN_IN);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_ITEM_CARDS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_ROTE_CARDS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_LOCATION_CARDS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_XP_HISTORY);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_PP_HISTORY);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_PLAYER_ROSTER);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_GAME_CALENDAR);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_SEARCH);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_STATISTICS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_MERITS_FLAWS);
        outputTemplates.add(placeholder);
        placeholder.setName(Constants.TN_INFLUENCE);
        outputTemplates.add(placeholder);
    }

    /**
     * Add default XP and PP awards to this game
     */
    private void addDefaultXPAwards() {
        String[] xpDefaults =
                {"Attendance","Costuming","Downtime Activities","First Night","Good Roleplaying","Leadership"};
        String[] ppDefaults = {"Storytelling", "Bookkeeping", "Narrating", "Setup/Cleanup"};
        ExperienceAward award;

        for (String xp : xpDefaults) {
            award = new ExperienceAward();
            award.setName(xp);
            award.setXp(true);
            award.setChange(ExperienceChange.EARNED);
            award.setAmount(1);
            award.setReason(xp);
            xpAwardList.add(award);
        }

        for (String pp : ppDefaults) {
            award = new ExperienceAward();
            award.setName(pp);
            award.setXp(false);
            award.setChange(ExperienceChange.EARNED);
            award.setAmount(1);
            award.setReason(pp);
            xpAwardList.add(award);
        }
    }

    /**
     * ToDo: Whatever needs done here when I actually implement a progress bar, if anything.
     */
    private void addFileProgress(int addition) { }

    /**
     * ToDo: Implement this?
     **/
    private void deleteDuplicates(List list, String warning) { }

    /**
     * ToDo: Implement this?
     **/
    private void ensureNoDuplicates() { }

    /**
     * ToDo: Implement this?
     **/
    private void entityCount(List list, String warning) { }

    /**
     * ToDo: Implement this?
     **/
    private void getValue(String key, Object value) { }

    /**
     * ToDo: Implement this?
     **/
    public void initializeForOutput() { }

    /**
     * ToDo: Implement this?!
     **/
    private void insertDuplicate(List list, String warning) { }

    /**
     * Merge an exchange file with the current game's data.
     * @param fileName the file to load
     * @throws IOException on error or invalid file
     */
    public void loadExchange(String fileName) throws IOException {
        ObjectInputStream inputStream = new ObjectInputStream(new FileInputStream(fileName));
        switch (ImportHelper.detectFileFormat(inputStream)) {
            case BINARY_EXCHANGE:
                loadExchangeBinary(inputStream);
            case XML:
                loadExchangeXML(inputStream);
            case BINARY_GAME:
            case BINARY_MENU:
            case INVALID:
                throw new IOException("File is not an Exchange file.");
        }
    }

    /**
     * Load a selection of data from a binary file. Should only be called when the file type has been ensured.
     * @param inputStream the input stream used to identify the file type.
     */
    private void loadExchangeBinary(ObjectInputStream inputStream) throws IOException{
        double version;
        version = inputStream.readDouble();
        if (version >= 2.395) {
            try {
                //ToDo: Decide if we're handing the calendar this way from here on in.
                GameCalendar newCalendar = GameCalendar.inputFromBinary(inputStream);
                boolean overwrite = false; //ToDo: Ask if the user wants to replace the current game calendar
                if (overwrite) {
                    this.gameCalendar = newCalendar;
                }
            } catch (ClassNotFoundException e) {
                throw new IOException("Malformed Exchange File (Calendar Break).", e);
            }
            if (version >= 2.397) {
                try {
                    APREngine newAPREngine = APREngine.inputFromBinary(inputStream);
                    boolean overwrite = false; //ToDo: Ask if the user wants to replace the current action/plot/rumor settings
                    if (overwrite) {
                        newAPREngine.setActions(this.aprEngine.getActions());
                        newAPREngine.setPlots(this.aprEngine.getPlots());
                        newAPREngine.setRumors(this.aprEngine.getRumors());
                        this.aprEngine = newAPREngine;
                    }
                } catch (ClassNotFoundException e) {
                    throw new IOException("Malformed Exchange File (APREngine Break)..", e);
                }
                try {
                    //ToDo: Consider asking here too.
                    int quantity = inputStream.readInt();
                    for (int i = 0; i < quantity; i++) {
                        ExperienceAward newAward = ExperienceAward.inputFromBinary(inputStream, version);
                        for (ExperienceAward award : xpAwardList) {
                            if (award.getName().equals(newAward.getName())) {
                                xpAwardList.remove(award);
                            }
                        }
                        xpAwardList.add(newAward);
                    }
                } catch (ClassNotFoundException e) {
                    throw new IOException("Malformed Exchange File (ExperienceAward Break)..", e);
                }
                try {
                    //ToDo: Consider asking here too.
                    int quantity = inputStream.readInt();
                    for (int i = 0; i < quantity; i++) {
                        Template newTemplate = Template.inputFromBinary(inputStream, version);
                        for (Template template : outputTemplates) {
                            if (template.getName().equals(newTemplate.getName())) {
                                outputTemplates.remove(template);
                            }
                        }
                        outputTemplates.add(newTemplate);
                    }
                } catch (ClassNotFoundException e) {
                    throw new IOException("Malformed Exchange File (Template Break)..", e);
                }
            }
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Player newPlayer = Player.inputFromBinary(inputStream, version);
                for (Player player : players) {
                    if (player.getName().equals(newPlayer.getName())) {
                        newPlayer = player.resolveDuplicate(newPlayer);
                        players.remove(player);
                    }
                }
                players.add(newPlayer);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Player Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            Character newCharacter;
            for (int i = 0; i < quantity; i++) {
                switch ((Race) inputStream.readObject()) {
                    case VAMPIRE:
                        newCharacter = Vampire.inputFromBinary(inputStream, version);
                        break;
                    case WEREWOLF:
                        newCharacter = Werewolf.inputFromBinary(inputStream, version);
                        break;
                    case MAGE:
                        newCharacter = Mage.inputFromBinary(inputStream, version);
                        break;
                    case CHANGELING:
                        newCharacter = Changeling.inputFromBinary(inputStream, version);
                        break;
                    case WRAITH:
                        newCharacter = Wraith.inputFromBinary(inputStream, version);
                        break;
                    case MORTAL:
                        newCharacter = Mortal.inputFromBinary(inputStream, version);
                        break;
                    case MUMMY:
                        newCharacter = Mummy.inputFromBinary(inputStream, version);
                        break;
                    case KUEIJIN:
                        newCharacter = KueiJin.inputFromBinary(inputStream, version);
                        break;
                    case FERA:
                        newCharacter = Fera.inputFromBinary(inputStream, version);
                        break;
                    case HUNTER:
                        newCharacter = Hunter.inputFromBinary(inputStream, version);
                        break;
                    case DEMON:
                        newCharacter = Demon.inputFromBinary(inputStream, version);
                        break;
                    case VARIOUS:
                    default:
                        newCharacter = Various.inputFromBinary(inputStream, version);
                        break;
                }
                for (Character character : characters) {
                    if (character.getRace().equals(newCharacter.getRace()) &&
                            character.getName().equals(newCharacter.getName())) {
                        newCharacter = character.resolveDuplicate(newCharacter);
                        players.remove(character);
                    }
                }
                characters.add(newCharacter);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Character Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Query newQuery = Query.inputFromBinary(inputStream, version);
                for (Query query : queryEngine.getQueryList()) {
                    if (query.getName().equals(newQuery.getName())) {
                        newQuery = query.resolveDuplicate(newQuery);
                        queryEngine.resolveDuplicateQuery(query, newQuery);
                    }
                }
                queryEngine.addQuery(newQuery);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Query Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Rote newRote = Rote.inputFromBinary(inputStream, version);
                for (Rote rote : rotes) {
                    if (rote.getName().equals(newRote.getName())) {
                        newRote = rote.resolveDuplicate(newRote);
                        rotes.remove(rote);
                    }
                }
                rotes.add(newRote);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Rote Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Location newLocation = Location.inputFromBinary(inputStream, version);
                for (Location location : locations) {
                    if (location.getName().equals(newLocation.getName())) {
                        newLocation = location.resolveDuplicate(newLocation);
                        locations.remove(location);
                    }
                }
                locations.add(newLocation);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Location Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Action newAction = Action.inputFromBinary(inputStream, version);
                for (Action action : aprEngine.getActions()) {
                    if (action.getName().equals(newAction.getName())) {
                        newAction = aprEngine.resolveDuplicateAction(action, newAction);
                        aprEngine.removeAction(action);
                    }
                }
                aprEngine.addAction(newAction);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Action Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Plot newPlot = Plot.inputFromBinary(inputStream, version);
                for (Plot plot : aprEngine.getPlots()) {
                    if (plot.getName().equals(newPlot.getName())) {
                        newPlot = aprEngine.resolveDuplicatePlot(plot, newPlot);
                        aprEngine.removePlot(plot);
                    }
                }
                aprEngine.addPlot(newPlot);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Plot Break)..", e);
        }
        try {
            //ToDo: Consider asking here too.
            int quantity = inputStream.readInt();
            for (int i = 0; i < quantity; i++) {
                Rumor newRumor = Rumor.inputFromBinary(inputStream, version);
                for (Rumor rumor : aprEngine.getRumors()) {
                    if (rumor.getName().equals(newRumor.getName())) {
                        newRumor = aprEngine.resolveDuplicateRumor(rumor, newRumor);
                       aprEngine.removeRumor(rumor);
                    }
                }
                aprEngine.addRumor(newRumor);
            }
        } catch (ClassNotFoundException e) {
            throw new IOException("Malformed Exchange File (Rumor Break)..", e);
        }
        changed = true;
        inputStream.close();
    }

    //ToDo: Marshalling and unmarshalling XML should go here.

    /**
     * ToDo: Implement this?!
     **/
    public void newGame() {
    }

    /**
     * ToDo: Implement this?!
     **/
    public void openGame() {
    }

    /**
     * ToDo: Implement this?!
     **/
    public void saveGame() {
    }


}