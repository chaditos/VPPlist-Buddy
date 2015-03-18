namespace VPPListBuddy.Workflow


type Alert = delegate of unit -> unit
type FileError = delegate of string -> unit
type DirectoryChooser = delegate of unit -> string
type PartitionWorkflowSetup = delegate of PartitionWorkflow -> unit