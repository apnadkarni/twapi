# MeTOO stands for "Metoo Emulates TclOO" (at a superficial syntactic level)
#
# Implements a *tiny* subset of TclOO, primarily for use with Tcl 8.4.
# Intent is that if you write code using MeToo, it should work unmodified
# with TclOO in 8.5/8.6. Obviously, don't try going the other way!
#
# Emulation is superficial, don't try to be too clever in usage.
# Renaming classes or objects will in all likelihood break stuff.
# Doing funky, or even non-funky, things with object namespaces will
# not work as you would expect.
#
# Note - MeTOO is class-based 

catch {namespace delete metoo}

namespace eval metoo {
    variable next_id 0
}

# Namespace in which commands in a class definition block are called
namespace eval metoo::define {
    proc method {class_ns name params body} {
        # Methods are defined in the methods subspace of the class namespace
        # The first parameter to a method is always the object namespace
        # denoted as the paramter "_this"
        namespace eval ${class_ns}::methods [list proc $name [concat [list _this] $params] $body]

    }
    proc superclass {class_ns superclass} {
        set ${class_ns}::super [uplevel 1 "namespace eval $class_ns {namespace current}"]
    }
}

# Namespace in which commands used in objects methods are defined
# (self, my etc.)
namespace eval metoo::object {
    proc self {{what object}} {
        upvar 1 _this this
        switch -exact -- $what {
            namespace { return $this }
            object { return [set ${this}::_(name)] }
            default {
                error "Argument '$what' not understood by self method"
            }
        }
    }

    proc my {methodname args} {
        # We insert the object namespace as the first parameter to the command.
        # This is passed as the first parameter "_this" to methods. Since
        # "my" can be only called from methods, we can retrieve it fro
        # our caller.
        upvar 1 _this this

        # We need to invoke in the caller's context so upvar etc. will
        # not be affected by this intermediate method dispatcher
        uplevel 1 [list [namespace parent $this]::methods::$methodname $this] $args
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

    # Create the namespace. The array _ is used to hold private information
    namespace eval $objns {
        variable _
    }
    set ${objns}::_(name) $objname

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

    # Define the built in commands callable within class instance methods
    foreach procname [info commands [namespace current]::object::*] {
        interp alias {} ${class_ns}::methods::[namespace tail $procname] {} $procname
    }

    # Define the class
    namespace eval $class_ns $definition

    # Also define the call dispatcher within the class. This is to get
    # the namespaces right when dispatching via "my"
    namespace eval ${class_ns} {
        proc _call {objns args} {
            set _this $objns;   # Proc my does an uplevel on _this to get objns
            eval [namespace current]::methods::my $args
        }
    }

    # The namespace is also a command used to create class instances
    interp alias {} $class_ns {} [namespace current]::_new $class_ns

    return $class_ns
}

