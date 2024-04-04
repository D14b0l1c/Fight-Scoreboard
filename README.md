These VBA modules collectively facilitate the management of a scoreboard and timers for a two-fighter scenario. The functionalities provided include:

1. Scoreboard Management:
   - Scoreboard modules for both fighters (`Takedown`, `Reversal`, `Escape`, `RunTime`, `Penalty`, and `PenaltyX`) are implemented. These modules update respective scores and log the actions with timestamps into designated worksheets.

2. Timer Control for Fighter One and Fighter Two:
   - For each fighter, there are modules to start, stop, and reset timers (`start_overtimer`, `stop_overtimer`, and `reset_overtimer`). These timers are designed to manage specific actions or tasks with a set duration.
   - Additionally, each fighter has its own ride timer management system (`RideTimer_Start`, `RideTimer_Stop`, and `RideTimer_Reset`). These timers allow tracking elapsed time with pause and reset functionalities.

3. Global Variables and Flags:
   - Global variables (`interval`, `StopIt`, `ResetIt`, `LastTime`) are utilized for maintaining timer states and managing actions.
   - Boolean flags (`StopIt`, `ResetIt`) are employed to control the start, stop, and reset functionalities of the timers.

4. Logging:
   - Logging functionalities are integrated into various scoreboard and timer modules to record actions and timestamps into corresponding worksheets for later analysis.

Overall, these VBA modules provide a comprehensive solution for managing scores and timers in a two-fighter scenario, offering precise control over actions, durations, and logging for effective event management.
