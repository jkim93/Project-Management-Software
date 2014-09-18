Project-Management-Software
===========================

Project Management Software within Microsoft Excel using VBA

Instructions: Open Project Management Software.xlsm to access the program

Summary:
  A full fledged project management software using Microsoft Excel and VBA.
•	Creates a representation of the Activity On Node (AON) network as a list with the predecessors and the successors for each activity.
•	Allows the modification of the list, and can save the modified list as a text file.
•	Converts AON to Activity On Arc (AOA) network with a representation similar to AON network.
•	Uses AOA network to build a linear program for finding earliest start (ES) and earliest finish (EF) times.
•	Uses the same linear program to find latest start (LS) and latest finish (LF) times.
•	Uses ES and LS (or EF and LF) to determine the critical activities and the critical path.
•	Reports the project completion time for deterministic project network.
•	Reports expected project completion time and the standard deviation for probabilistic project network.
•	Computes probability of project completion within a user‐defined time period.
•	Creates a Gantt chart based on the solution and highlights the critical activities.
•	Uses AOA network to build a linear program for crashing the project down to a user‐defined time period at minimum cost; finds the solution that maximizes project crash amount at minimum cost.
•	Creates a Gantt chart based on the crashed solution and highlights the crashed activities and the critical activities, displays crash amount and crash cost for each activity.
•	You can consider implementing the following functionalities as bonus for up to an additional 10% credit.
•	Forward pass and backward pass algorithms to find the ES, EF, LS, and LF to find the critical activities and the critical path as an alternative to the linear programming approach.
