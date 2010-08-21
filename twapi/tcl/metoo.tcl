# metoo stands for "Metoo Emulates TclOO"
#
# Implements a *tiny* subst of TclOO, primarily for use with Tcl 8.4.
# Intent is that if you write code using MeToo, it should work unmodified
# with TclOO in 8.5/8.6. Obviously, don't try going the other way!
#
# Emulation is superficial, don't try to be too clever in usage.
# Renaming classes or objects will in all likelihood break stuff.

catch {namespace delete metoo}

namespace eval metoo {
    variable next_id 0
}

# Namespace in which commands in a class definition block are called
namespace eval metoo::define {
    proc method {class_ns name params body} {
        # The first parameter to a method is always the object namespace
        proc ${class_ns}::$name [concat [list _this] $params] $body
    }
    proc superclass {class_ns superclass} {
        set ${class_ns}::super [uplevel 1 "namespace eval $class_ns {namespace current}"]
    }
}

# Namespace in which commands used in objects methods are defined
# (self, my etc.)
namespace eval metoo::object {
    proc self {obj_ns {what object}} {
        switch -exact -- $what {
            namespace { return $obj_ns }
            object { return ${obj_ns}::_(name) }
            default {
                error "Argument '$what' not understood by self method"
            }
        }
    }

    proc my {methodname args} {
        # Class namespace is always parent of object namespace
        # We insert the object namespace as the first parameter to the command
        # We need to invoke in the caller's context so upvar etc. will
        # not be affected by this intermediate method dispatcher
        upvar 1 _this this
        uplevel 1 [list $methodname $this] $args
    }
}

proc metoo::_new {class_ns cmd args} {
    variable next_id

    # object namespace *must* be child of class namespace. Saves a bit
    # of bookkeeping
    set objns ${class_ns}::[incr next_id]

    switch -exact -- $cmd {
        create {
            if {[llength $args] < 1} {
                error "Insufficient args, should be: class create CLASSNAME ?args?"
            }
            # TBD - check if command already exists
            set objname [uplevel 1 "namespace current"]::[lindex $args 0]
            set args [lrange $args 1 end]
        }
        new {
            set objname $objns
        }
        default {
            error "Unknown command '$cmd'. Should be create or new."
        }
    }

    # Create the namespace. Add built-in commands to it
    namespace eval $objns {}

    # Invoke the constructor
    # TBD

    # When invoked by its name, call the dispatcher
    interp alias {} $objname {} ${class_ns}::_call $objns

    return $objname
}


proc metoo::class {cmd cname definition} {
    variable next_id

    if {$cmd ne "create"} {
        error "Syntax: class create CLASSNAME DEFINITION"
    }

    if {[uplevel 1 "namespace exists $cname"]} {
        error "can't create class '$cname': namespace already exists with that name."
    }

    # Resolve cname into a namespace in the caller's context
    set class_ns [uplevel 1 "namespace eval $cname {namespace current}"]
    
    if {[llength [info commands $class_ns]]} {
        # Delete the namespace we just created
        namespace delete $class_ns
        error "can't create class '$cname': command already exists with that name."
    }

    # Define the commands/aliases that are used inside a class definition
    foreach procname [info commands [namespace current]::define::*] {
        interp alias {} ${class_ns}::[namespace tail $procname] {} $procname $class_ns
    }

    foreach procname [info commands [namespace current]::object::*] {
        interp alias {} ${class_ns}::[namespace tail $procname] {} $procname
    }

    # Define the class
    namespace eval $class_ns $definition

    # Also define the call dispatcher within the class. This is to get
    # the namespaces right when dispatching via "my"
    namespace eval $class_ns {
        proc _call {objns args} {
            set _this $objns;   # Proc my does an uplevel on _this to get objns
            eval [list [namespace current]::my] $args
        }
    }

    # The namespace is also a command used to create class instances
    interp alias {} $class_ns {} [namespace current]::_new $class_ns

    return $class_ns
}

