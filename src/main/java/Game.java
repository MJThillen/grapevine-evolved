import constants.FileFormat;
import constants.Race;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class Game {
    private List players; //Collection of all players
    private List characters; //collection of all characters
    private List items; //collection of all items
    private List rotes; //collection of all rotes
    private List locations; //collection of all locations
    private List rumorLists; //collection of lists of rumors, one per date
    private List influenceUses; //collection of influence use actions
    private MenuSet menuSet; //collection of menus
    private QueryEngine queryEngine; //collection, logic of queries
    private APREngine aprEngine; //collection, logic of actions/plots/rumors

    private FileFormat fileFormat; //current file type to save file as
    private String gameFilename; //full pathname of game file

    private String chronicleTitle; //Title of the Chronicle
    private String website; //Game URL
    private String email; //Main ST Email address
    private String phone; //Main Phone Number
    private String usualSite; //Usual Game Site
    private String usualTime; //Usual Game Start Time
    private Calendar calendar; //Calendar of game dates
    private String description; //Description of your game
    private boolean extendedHealth; //whether this game uses abbreviated or extended health levels
    private boolean enforceHistory; //whether to enforce use of XP history
    private boolean linkTraitMaxes; //Whether to link trait maximums on character sheets
    private String randomTraits; //Comma-separated list of random trait options
    private String stCommentStart; //opening markup of an ST comment
    private String stCommentEnd; //closing markup of an ST comment
    private List xpAwardList; //list of standard XP and PP awards
    private List outputTemplates; //list of output templates
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
        this.calendar = Calendar.getInstance();
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
     * Name: "Add Character from File"
     * Create a new character object of the appropriate type and have it load itself from the files.
     * Used for v2.3 files only -- this method survives for backwards compatibility.
     *
     * per VB6: Required GUI class: frmDuplicate
     * @param race      a string representation of the character's race
     * @param fileNum   an open file number
     * @param oldDate   the date of the old file
     * @param version   the file format version tag
     * @return The new character loaded
     * @throws Exception if the character type is a mismatch to the Enum.
     */
    private Character addV2Character(final String race,
                                 final int fileNum,
                                 final Date oldDate,
                                 final String version) throws Exception {
        Character newCharacter;
        String newName;

        Race raceType = Race.valueOf(race.trim());

        switch(raceType) {
            case VAMPIRE:
                newCharacter = new Vampire();
                break;
            case WEREWOLF:
                newCharacter = new Werewolf();
                break;
            case MORTAL:
                newCharacter = new Mortal();
                break;
            case CHANGELING:
                newCharacter = new Changeling();
                break;
            case WRAITH:
                newCharacter = new Wraith();
                break;
            case MAGE:
                newCharacter = new Mage();
                break;
            case FERA:
                newCharacter = new Fera();
                break;
            case VARIOUS:
                newCharacter = new Various();
                break;
            case MUMMY:
                newCharacter = new Mummy();
                break;
            case KUEIJIN:
                newCharacter = new KueiJin();
                break;
            case HUNTER:
                newCharacter = new Hunter();
                break;
            case DEMON:
                newCharacter = new Demon();
                break;
            default:
                newCharacter = null;
                throw new Exception("Unexpected Character Type: " + race);
        }
        newCharacter.oldInputFromFile(fileNum, oldDate, version);
        if (characters.contains(newCharacter)) {
            insertDuplicate(newCharacter);
        }
        else {
            characters.add(newCharacter);
        }
        return newCharacter;
    }

    /**
     * Adds the default output templates to this game.
     */
    private void addDefaultTemplates() {

    }
}