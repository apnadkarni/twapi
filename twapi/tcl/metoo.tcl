# MeTOO stands for "MeTOO Emulates TclOO" (at a superficial syntactic level)
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
# Differences:
# - MeTOO is class-based, not object based like TclOO, thus class instances
#   (objects) cannot be modified. Also a class is not itself an object
# - does not support class refinement/definition. This would not actually
#   be hard to add but I wanted to keep this small
# - The unknown method is not supported. Again should not be hard to add
# - no filters, forwarding, multiple-inheritance
# - no introspection capabilities


catch {namespace delete metoo}

# TBD - put a trace to delete object when the command is deleted
# TBD - put a trace when command is renamed
# TBD - delete all objects when a class is deleted
# TBD - delete all subclasses when a class is deleted
# TBD - implement next
# TBD - implement destructor
# TBD - variable

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
        set ${class_ns}::super [uplevel 3 "namespace eval $superclass {namespace current}"]
    }
    proc constructor {class_ns params body} {
        method $class_ns constructor $params $body
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
        upvar 1 _this this;     # object namespace

        set class_ns [namespace parent $this]

        # See if there is a method defined in this class.
        # Breakage if method names with with wildcard chars. Too bad
        if {[llength [info commands ${class_ns}::methods::$methodname]]} {
            # We need to invoke in the caller's context so upvar etc. will
            # not be affected by this intermediate method dispatcher
            return [uplevel 1 [list ${class_ns}::methods::$methodname $this] $args]
        }

        # No method here, check for super class.
        while {[info exists ${class_ns}::super]} {
            set class_ns [set ${class_ns}::super]
            if {[llength [info commands ${class_ns}::methods::$methodname]]} {
                return [uplevel 1 [list ${class_ns}::methods::$methodname $this] $args]
                # return [eval [list [set ${class_ns}::super]::_call $this $methodname] $args]

            }
        }

        # It is ok for constructor or destructor to be undefined
        if {$methodname ne "constructor" && $methodname ne "destructor"} {
            error "Unknown method $methodname"
        }
        return
    }
}

proc metoo::_new {class_ns cmd args} {
    variable next_id

    # object namespace *must* be child of class namespace. Saves a bit
    # of bookkeeping
    set objns ${class_ns}::o[incr next_id]

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

    # When invoked by its name, call the dispatcher
    interp alias {} $objname {} ${class_ns}::_call $objns

    # Invoke the constructor
    eval [list $objname constructor] $args

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

