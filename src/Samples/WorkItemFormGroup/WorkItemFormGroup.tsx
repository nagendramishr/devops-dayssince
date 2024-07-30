import {
  IWorkItemChangedArgs,
  IWorkItemFieldChangedArgs,
  IWorkItemFormService,
  IWorkItemLoadedArgs,
  WorkItemTrackingServiceIds
} from "azure-devops-extension-api/WorkItemTracking";
import * as SDK from "azure-devops-extension-sdk";
import { Button } from "azure-devops-ui/Button";
import * as React from "react";
import { showRootComponent } from "../../Common";

interface WorkItemFormGroupComponentState {
  eventContent: string;
}

export class WorkItemFormGroupComponent extends React.Component<{},  WorkItemFormGroupComponentState> {
  constructor(props: {}) {
    super(props);
    this.state = {
      eventContent: ""
    };
  }

  public componentDidMount() {
    SDK.init().then(() => {
      this.registerEvents();
    });
  }

  public render(): JSX.Element {
    return ( <div className="sample-work-item-events">{this.state.eventContent}</div> );
  }

  private registerEvents() {
    SDK.register(SDK.getContributionId(), () => {
      return {

        // onSaved: async (args: IWorkItemFieldChangedArgs) => {
        //   const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);

        //   workItemFormService.setFieldValue( "Custom.ModifiedDate", new Date().toISOString() ).then(() => { 
        //     workItemFormService.isDirty().then(isDirty => { 
        //       //console.log('isDirty: ' + isDirty)
        //       if (isDirty) { 
        //         workItemFormService.save().then(() => { 
        //         //console.log("Saved"); 
        //         console.log('ModifiedDate field updated to : ' + new Date().toISOString() ) 
        //         });
        //       }
        //     });
        //   });          
        // },

        onFieldChanged: async (args: IWorkItemFieldChangedArgs) => {
          const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
        
          if (args.changedFields ) {
            const fieldsToCheck = ["System.Title", "System.Description", "System.AreaPath", "System.IterationPath", "System.Tags", 
            //"System.WorkItemType", "System.AssignedTo", "System.History"];
            "System.History"];
          
            for (let field in args.changedFields) {
              if ((field.startsWith("Custom") && ( field !== "Custom.Age" && field != "Custom.ModifiedDate" )) ||
               (field.startsWith("Microsoft") && field !="Microsoft.VSTS.Scheduling.StartDate") || 
               fieldsToCheck.includes(field)) {
         
                await workItemFormService.setFieldValue("Custom.ModifiedDate", new Date().toISOString());
                console.log("Updated ModifiedDate field because of change in " + field);
                break;
              }
              console.log(JSON.stringify(args));
            }

            for (let field in args.changedFields) {
              if (field == "Microsoft.VSTS.Scheduling.StartDate") {
                const startDate = await workItemFormService.getFieldValue("Microsoft.VSTS.Scheduling.StartDate");
                const age= Math.floor((( new Date().getTime() - new Date(startDate.toString()).getTime() ) / 86400000)); // 86400000=1000*60*68*24 = number of milliseconds/day

                this.setState({ eventContent:  age + ' days' });
              }
            }

          }
        },

      
        // Called when a new work item is being loaded in the UI
        onLoaded: async (args: IWorkItemLoadedArgs) => {

          const workItemFormService = await SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService);
          let age = 0;

          workItemFormService.getFieldValue("Custom.Age").then(customAge => {
            age = parseInt(customAge.toString());  // This is the current value of age. If the status is closed, we will not update this value.
            
            workItemFormService.getFieldValue("System.State").then(state => {
            if (state != "Closed") {
            
              workItemFormService.getFieldValue("Microsoft.VSTS.Scheduling.StartDate").then(fields => {
                age= Math.floor((( new Date().getTime() - new Date(fields.toString()).getTime() ) / 86400000)); // 86400000=1000*60*68*24 = number of milliseconds/day

                workItemFormService.setFieldValue( "Custom.Age", age ).then(() => { 
                  //console.log('Age field updated to : ' + age) 
                  workItemFormService.isDirty().then(isDirty => { 
                    //console.log('isDirty: ' + isDirty)
                    if (isDirty) { 
                      workItemFormService.save().then(() => { 
                      //console.log("Saved"); 
                      } );
                    } else {
                     // workItemFormService.getFields().then(fields => {
                       //console.log("No need to save: " + JSON.stringify(fields) );
                       // } );
                    } // if isDirty
                  }); // isDirty
                }); //SetFieldValue - Custom.Age
              }); //GetFieldValue - CreatedDate
            } // if status != Closed
                        
//            if (state != "New") 
              this.setState({ eventContent:  age + ' days' });
//            else
//              this.setState({ eventContent:  age + ' days  - <i>NEW</i>' });
            //console.log( "Calculated Age: " + age + ' days' );
          }) // GetFieldValue - State
        });
        }, //onLoaded

      };
    });
  }
}

export default WorkItemFormGroupComponent;

showRootComponent(<WorkItemFormGroupComponent />);
