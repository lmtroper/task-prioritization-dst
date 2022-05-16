## Decision Support Tool For Task Prioritization (VBA, Excel)

The decision support tool was designed to aid with a client’s challenges of task prioritization and time-management with regards to university project planning and execution.

The tool will be able to sort the tasks from highest priority to lowest priority. Using prioritization sorting algorithm, the user will be able to view which tasks should be completed first - assisting the user with the efficient delegation of tasks. The prioritization algorithm utilizes subjective data obtained from the user, and thus, will not produce the same results for all users.

Additional goals of the tool are to provide the user with visual cues of upcoming or overdue tasks, the ability to filter tasks by deadline or assigned team members, and the ability to alter task data (i.e. edit, delete, complete a task).

#### Menu/Home Page
User can the menu page to navigate to the different sheets within the Excel program
![Menu](https://user-images.githubusercontent.com/72565412/168627295-57f2290b-9360-4d2c-a41d-0f5f7a8533cb.jpg)

### Main Task Sheet
User can add new tasks, update ongoing tasks, complete or delete tasks
![TaskSheet_Page](https://user-images.githubusercontent.com/72565412/168627683-8545f70f-5c5a-4e57-b1d6-62c1745ba9d3.jpg)

#### Add Task User Form
The user is prompted with data fields to fill out regarding the new task. Using information about the task's estimated time to complete, difficulty, and importance to the project, a prioritization algorithm sorts the tasks on the Task Sheet from highest to lowest priority.
![Add_Task_UserForm](https://user-images.githubusercontent.com/72565412/168627720-c638c5ad-daec-400f-90f7-72e47a0aec17.jpg)

#### Prioritization Algorithm
The prioritization algorithm was inspired using two existing task prioritization matrices:
<img width="638" alt="Screen Shot 2022-05-16 at 11 48 01 AM" src="https://user-images.githubusercontent.com/72565412/168632692-118c7bd6-a01a-45bc-92f1-3117d0df709e.png">

The principles taken from these two matrices are that tasks should be assigned a high priority ranking if they are considered important to the user and need to be completed urgently, and the consideration of how much effort a task needs in comparison to its importance. More specifically, if a task is not important yet requires great effort, it should be less prioritized than tasks of greater importance that require either great effort or little effort.

To remain consistent with the principles of the Action Priority matrix, when importance is less than 3, the prioritization equation takes the reciprocal of the effort of value. This ensures that less important tasks that require great effort are given a lower priority value. In contrast, more important tasks will use the unmodified value of effort in the prioritization algorithm. This ensures that tasks that are important to the user that require high effort will be prioritized high as the user will need to begin working on that task sooner than important tasks that require little effort.

For the algorithm, urgency is equal to difference of days between the deadline and current date; effort is determined by the weighted average of task difficulty and estimated time to complete, both values subjectively inputted by the user; and impact is grouped under the term importance, which is determined by the user using a subjective ranking system (1, being of low importance, to 5, being of high importance).

 ![image](https://user-images.githubusercontent.com/72565412/168632973-4d62460b-9795-4bf5-81a2-643750d9f3c2.png)


#### Edit Task User Form
The Edit Task Module is designed to provide the user an opportunity to edit the data about a task after it has been added.

User is prompted to select the task they would like to edit/update
![EditTask_UserForm2](https://user-images.githubusercontent.com/72565412/168627857-18a36d32-45c7-421f-b7ee-b67ef028be2e.jpg)

User is then presented with the current information associated with the selected task. When data is edited, the Task Sheet is updated with the new information. Editing the data could potentially alter the priority ranking - thus, the prioritization algorithm is called, and the task is assigned a new priority ranking. The ranking may be the same if the data inputted is unrelated to the prioritization of the task (i.e. task name, assigned team member, category). 

![EditTask_UserForm1](https://user-images.githubusercontent.com/72565412/168627869-48b384ad-31ab-4ebc-960b-db3ef26293bf.jpg)

### View Tasks Sheet
The user can filter tasks from a date range or view them by a specific team member. This mitigates the inconvenience of being unable to see desired information due to a large volume of entries on the TaskSheet.

![ViewTask_Sheet](https://user-images.githubusercontent.com/72565412/168627978-67052275-612c-4c33-ba25-ad1e28529a5e.jpg)

#### User Form to Filter and View Ongoing Tasks
![ViewTasks_UserForm](https://user-images.githubusercontent.com/72565412/168628763-8c74c257-3d1d-4126-96a6-80ae9f44a859.jpg)

### Completed Tasks Sheet
The Completed Task Module is designed to remove the user’s completed tasks from the Task Sheet. When tasks are marked as completed, they are transferred to the Completed Tasks Sheet. 


Here, users can view and keep a log of tasks they have completed.
![CompletedTasks_Sheet](https://user-images.githubusercontent.com/72565412/168628042-9e142513-7be1-47f8-9c8d-cb2b6607e89d.jpg)

#### User Form to Mark the Task as Completed
User can select which tasks they would like to mark as completed
![CompleteTask_UserForm](https://user-images.githubusercontent.com/72565412/168628029-d236edea-5862-402e-b5b0-fc929bc0d801.jpg)

Users are prompted to confirm the changes to be made.
![CompleteTask_UserForm2](https://user-images.githubusercontent.com/72565412/168628037-66f101f0-d72b-421e-9f9f-8d8072031e96.jpg)
