# stgvfs.tcl - Copyright (C) 2004 Pat Thoyts <patthoyts@users.sourceforge.net>
#
# 	This maps an OLE Structured Storage into a Tcl virtual filesystem
#	You can mount a DOC file (word or excel or any other such compound
#	document) and then treat it as a filesystem. All sub-storages are
#	presented as directories and all streams as files and can be opened
#	as tcl channels.
#	This vfs does not provide access to the property sets.
#

package require vfs 1;                  # tclvfs
package require Storage;                # tclstorage

namespace eval ::vfs::stg {
    variable version 1.0.0
    variable rcsid {$Id$}

    variable uid
    if {![info exists uid]} {
        set uid 0
    }
}

proc ::vfs::stg::Mount {path local} {
    variable uid
    set mode r+
    if {$path eq {}} { set mode w+ }
    set stg [::storage open [::file normalize $path] $mode]

    set token [namespace current]::mount[incr uid]
    upvar #0 $token state
    catch {unset state}
    set state(/stg) $stg
    set state(/root) $path
    set state(/mnt) $local

    vfs::filesystem mount $local [list [namespace origin handler] $token]
    vfs::RegisterMount $local [list [namespace origin Unmount] $token]
    return $token
}

proc ::vfs::stg::Unmount {token local} {
    upvar #0 $token state
    foreach path [array get state] {
        if {![string match "/*" $path]} {
            catch {$state($path) close}
        }
    }
    vfs::filesystem unmount $local
    $state(/stg) close
    unset state
}

proc ::vfs::stg::Execute {path} {
    Mount $path $path
    source [file join $path main.tcl]
}

# -------------------------------------------------------------------------

proc ::vfs::stg::handler {token cmd root relative actualpath args} {
    #::vfs::log [list $token $cmd $root $relative $actualpath $args]
    if {![string compare $cmd "matchindirectory"]} {
	eval [linsert $args 0 $cmd $token $relative $actualpath]
    } else {
	eval [linsert $args 0 $cmd $token $relative]
    }
}

# Open up a path within the specified storage. We cache the intermediate
# opened storage items (we have to or the leaves are invalid).
# Returns the final storage item in the path.
# - path
proc ::vfs::stg::PathToStg {token path} {
    upvar #0 $token state
    set stg $state(/stg)
    if {[string equal $path "."]} { return $stg }
    if {[info exists state($path)]} {
        return $state($path)
    }
    if {[info exists state([file dirname $path])]} {
        return $state([file dirname $path])
    }

    set elements [file split $path]
    set path {}
    foreach dir $elements {
        set path [file join $path $dir]
        if {[info exists state($path)]} {
            set stg $state($path)
        } else {
            set stg [$stg opendir $dir r+]
            set state($path) $stg
        }
    }

    return $stg
}
    

# -------------------------------------------------------------------------
# The vfs handler procedures
# -------------------------------------------------------------------------

proc vfs::stg::access {token path mode} {
    ::vfs::log "access: $token $path $mode"

    if {[catch {
        if {[string length $path] < 1} { return }
        set stg [PathToStg $token [file dirname $path]]
        ::vfs::log "access: check [file tail $path] within $stg"
        if {[catch {$stg stat [file tail $path] sd} err]} {
            ::vfs::log "access: error: $err"
            ::vfs::filesystem posixerror $::vfs::posix(ENOENT)
        } else {
            if {($mode & 2) && !($sd(mode) & 2)} {
                ::vfs::filesystem posixerror $::vfs::posix(EACCES)
            }
        }
    } err]} { ::vfs::log "access: error: $err" }
    return
}

proc vfs::stg::createdirectory {token path} {
    ::vfs::log "createdirectory: $token \"$path\""
    upvar #0 $token state
    set stg [PathToStg $token [file dirname $path]]
    set state($path) [$stg opendir [file tail $path] w+]
}

proc vfs::stg::stat {token path} {
    ::vfs::log "stat: $token \"$path\""
    set stg [PathToStg $token [file dirname $path]]
    $stg stat [file tail $path] sb
    array get sb
}

proc vfs::stg::matchindirectory {token path actualpath pattern type} {
    ::vfs::log [list matchindirectory: $token $path $actualpath $pattern $type]

    set names {}
    if {[string length $pattern] > 0} {
        set stg [PathToStg $token $path]
        foreach name [$stg names] {
            if {[string match $pattern $name]} {lappend names $name}
        }
    } else {
        set stg [PathToStg $token [file dirname $path]]
        set names [file tail $path]
        set actualpath [file dirname $actualpath]
        if {[catch {$stg stat $names sd}]} {
            ::vfs::filesystem posixerror ::vfs::posix(ENOENT)
            return {}
        }
    }

    set glob {}
    foreach name [::vfs::matchCorrectTypes $type $names $actualpath] {
	lappend glob [file join $actualpath $name]
    }
    return $glob
}

proc vfs::stg::open {token path mode permissions} {
    ::vfs::log "open: $token \"$path\" $mode $permissions"
    set stg [PathToStg $token [file dirname $path]]
    if {[catch {set f [$stg open [file tail $path] $mode]} err]} {
        vfs::filesystem posixerror $::vfs::posix(EACCES)
    } else {
        return [list $f]
    }
}

proc vfs::stg::removedirectory {token path recursive} {
    ::vfs::log "removedirectory: $token \"$path\" $recursive"
    upvar #0 $token state
    set stg [PathToStg $token [file dirname $path]]
    $stg remove [file tail $path]
    if {[info exist state($path)]} {
        $state($path) close
        unset state($path)
    }
}

proc ::vfs::stg::deletefile {token path} {
    ::vfs::log "deletefile: $token \"$path\""
    set stg [PathToStg $token [file dirname $path]]
    $stg remove [file tail $path]
}

proc ::vfs::stg::fileattributes {token path args} {
    ::vfs::log "fileattributes: $token \"$path\" $args"
    # We don't have any yet.
    # We could show the guid and state fields from STATSTG possibly.
    switch -- [llength $args] {
	0 { return [list] }
	1 { return "" }
	2 { vfs::filesystem posixerror $::vfs::posix(EROFS) }
    }
}

proc ::vfs::stg::utime {token path atime mtime} {
    ::vfs::log "utime: $token \"$path\" $atime $mtime"
    set stg [PathToStg $token [file dirname $path]]
    #$stg touch [file tail $path] $atime $mtime
    # FIX ME: we don't have a touch op yet.
    vfs::filesystem posixerror $::vfs::posix(EACCES)
}

# -------------------------------------------------------------------------

package provide vfs::stg $::vfs::stg::version
package provide stgvfs   $::vfs::stg::version

# -------------------------------------------------------------------------
# Local variables:
#   mode: tcl
#   indent-tabs-mode: nil
# End:
