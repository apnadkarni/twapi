#ifndef ZLIST_H
#define ZLIST_H

/*
 * Set of macros of doubly linked lists. What's special? Nothing really.
 * The verbosity of definitions comes from two features -
 *  - the links/lists are type safe so you cannot point a link
 *    to a struct of the wrong type (unless of course you intentionally
 *    break the typing with a cast)
 *  - The links do not need to be the first element in the struct as is the
 *    case with many list macro packages out there.
 *
 * Lists are null terminated.
 *
 * Be *VERY* careful when assigning list items to each other since you
 * overwrite link fields resulting in cross-linked chains and general chaos.

 * Internal note: In all macros, use newobjptr_ fields as much as possible
 * in manipulating since curobjptr_ may change during manipulation.
 */


/*c
List handling
*/

#ifndef ZLIST_ASSERT
#ifndef NDEBUG
#define ZLIST_ASSERT(cond_) assert(cond_)
#else
#define ZLIST_ASSERT(cond_) (void) 0
#endif
#endif /* ZLIST_ASSERT */

/*
Declare a link element named field_ to link objects of type objtype_
*/
#define ZLINK_CREATE_TYPEDEFS(objtype_) \
    typedef struct { objtype_ *zl_prev; objtype_ *zl_next; } \
    ZLINK_TYPE(objtype_)

#define ZLINK_TYPE(objtype_)    zlink_ ## objtype_ ## _t

#define ZLINK_NAMED_DECL(objtype_, field_)     ZLINK_TYPE(objtype_) field_

/*
Initialize a link
*/
#define ZLINK_NAMED_INIT(objptr_, field_) \
    do { \
	(objptr_)->field_.zl_next = (objptr_)->field_.zl_prev = 0; \
     } while (0)

/*
Get the next/previous element in the list. Returns NULL if first/last element
*/
#define ZLINK_NAMED_NEXT(objptr_, field_) ((objptr_)->field_.zl_next)
#define ZLINK_NAMED_PREV(objptr_, field_) ((objptr_)->field_.zl_prev)

#define	ZLINK_NAMED_POSTATTACH(curobjptr_, newobjptr_, field_) \
    do { \
	/* Be careful if changing order assignments */ \
	(newobjptr_)->field_.zl_prev = (curobjptr_);  \
	(newobjptr_)->field_.zl_next = (curobjptr_)->field_.zl_next; \
	(curobjptr_)->field_.zl_next = (newobjptr_); \
	if ((newobjptr_)->field_.zl_next) \
	    (newobjptr_)->field_.zl_next->field_.zl_prev = (newobjptr_); \
    } while (0)

#define	ZLINK_NAMED_PREATTACH(curobjptr_, newobjptr_, field_) \
    do { \
	/* Be careful if changing order assignments */ \
	(newobjptr_)->field_.zl_next = (curobjptr_); \
	(newobjptr_)->field_.zl_prev = (curobjptr_)->field_.zl_prev; \
	(curobjptr_)->field_.zl_prev = (newobjptr_); \
	if ((newobjptr_)->field_.zl_prev) \
	    (newobjptr_)->field_.zl_prev->field_.zl_next = (newobjptr_); \
    } while (0)


#define ZLINK_NAMED_UNLINK(objptr_, field_) \
    do { \
	if ((objptr_)->field_.zl_prev) \
	    (objptr_)->field_.zl_prev->field_.zl_next = \
		(objptr_)->field_.zl_next; \
	if ((objptr_)->field_.zl_next) \
	    (objptr_)->field_.zl_next->field_.zl_prev = \
		(objptr_)->field_.zl_prev; \
	(objptr_)->field_.zl_next = (objptr_)->field_.zl_prev = 0; \
    } while (0)

#define ZLINK_NAMED_MOVETOFRONT(headptr_, objptr_, field_) \
    do { \
	if ((headptr_) != (objptr_)) { \
	    /* zl_prev guaranteed non-null */ \
	    (objptr_)->field_.zl_prev->field_.zl_next = \
		(objptr_)->field_.zl_next; \
	    if ((objptr_)->field_.zl_next) \
		(objptr_)->field_.zl_next->field_.zl_prev = \
		    (objptr_)->field_.zl_prev; \
	    (headptr_)->field_.zl_prev = (objptr_); \
	    (objptr_)->field_.zl_next = (headptr_); \
	    (objptr_)->field_.zl_prev = 0; \
	} \
    } while (0)

#define ZLINK_NAMED_SPLIT(objptr_, field_) \
    do { \
	if ((objptr_)->field_.zl_prev) \
	    (objptr_)->field_.zl_prev->field_.zl_next = 0; \
	(objptr_)->field_.zl_prev = 0; \
    } while (0)



/* Modifies objptr_ to point to first element in list */
#define ZLINK_NAMED_FIRST(objptr_, field_)  \
    while (ZLINK_NAMED_PREV(objptr_, field_)) { \
	objptr_ = ZLINK_NAMED_PREV(objptr_, field_); \
    }

/* Modifies objptr_ to point to last element in list */
#define ZLINK_NAMED_LAST(objptr_, field_)  \
    while (ZLINK_NAMED_NEXT(objptr_, field_)) { \
	objptr_ = ZLINK_NAMED_NEXT(objptr_, field_); \
    }

/*
Searches down the list, starting at objptr_ executing cmpmacro_ which must
be a macro or function takes two parameters - the first is the list element
being examined, the second is cmpmacroarg_ which is passed through unchanged.
The macro should return 0 if match succeeds, else non-0.
objptr_ will be set to the first element for which the match succeeded or
to null if non match occurred.
*/
#define ZLINK_NAMED_FIND(objptr_, field_, cmpmacro_, cmpmacroarg_) \
    do { \
	while (objptr_) { \
	    if (0 == cmpmacro_((objptr_), (cmpmacroarg_))) \
		break; \
	    (objptr_) = ZLINK_NAMED_NEXT((objptr_), field_); \
	} \
    } while (0)



/* Note that *tailPtrPtr is only used for returning the tail of the sorted
   list. The list that is to be sorted is assumed terminated with a null ptr */
#define ZLINK_NAMED_SORT_DECL(sortfunc_, objtypedef_, field_) \
    void sortfunc_ (objtypedef_ * *headPtrPtr, \
		    objtypedef_ * *tailPtrPtr, \
		    int (APNCALLBACK *cmpfunc_)(objtypedef_ *, \
				    objtypedef_ *))

#define ZLINK_NAMED_SORT_DEFINE(sortfunc_, objtypedef_, field_) \
    ZLINK_NAMED_SORT_DECL(sortfunc_, objtypedef_, field_) \
{ \
    /* This is based on a public domain quicksort routine to be used to */ \
    /* sort by Jon Guthrie. */ \
    objtypedef_ *current; \
    objtypedef_ *pivot; \
    objtypedef_ *temp; \
    int     result; \
    objtypedef_ *headPtr = *headPtrPtr; \
    objtypedef_ *tailPtr; \
    objtypedef_ *low; \
    objtypedef_ *lowTail; \
    objtypedef_ *high; \
    objtypedef_ *highTail; \
 \
    if (tailPtrPtr == 0) \
	tailPtrPtr = &tailPtr; \
 \
    if (headPtr == 0) \
	return; \
 \
    /* Find the <first element that doesn't have the same value as the first */ \
    current = headPtr; \
    do { \
	current = ZLINK_NAMED_NEXT(current, field_); \
	if (current == 0) \
	    return; \
    }   while(0 == (result = cmpfunc_(headPtr, current))); \
 \
    /*  pivot value is the lower of the two.  This insures that the sort */ \
    /*  will always terminate by guaranteeing that there will be at least*/ \
    /*  one member of both of the sublists. */ \
    if (result > 0) \
        pivot = current; \
    else \
        pivot = headPtr; \
 \
    /* Initialize the sublist pointers */ \
    low = lowTail = high = highTail = 0; \
 \
    /* Now, separate the items into the two sublists */ \
    current = headPtr; \
    while (current) { \
	temp = ZLINK_NAMED_NEXT(current, field_); \
	ZLINK_NAMED_UNLINK(current, field_); \
	if(cmpfunc_(pivot, current) < 0) { \
	    /* add one to the high list */ \
	    ZLINK_NAMED_PREATTACH(high, current, field_); \
	} \
	else { \
	    /* add one to the low list */ \
	    ZLINK_NAMED_PREATTACH(low, current, field_); \
	} \
	current = temp; \
    } \
 \
    /* And, recursively call the sort for each of the two sublists. */ \
    sortfunc_(&low, &lowTail, cmpfunc_); \
    sortfunc_(&high, &highTail, cmpfunc_); \
 \
    /* put the "high" list after the end of the "low" list. */ \
    ZLINK_NAMED_POSTATTACH(high, lowTail, field_); \
    *headPtrPtr = low; \
    *tailPtrPtr = highTail; \
    return; \
}


#define ZLIST_NAMED_CREATE_TYPEDEFS(objtypedef_, field_) \
    typedef struct { \
	objtypedef_ *zlist_head; \
	objtypedef_ *zlist_tail; \
	int zlist_count; \
    } ZLIST_NAMED_DECL(objtypedef_, field_)


/* Declare list of objects of type objtype_ linked through field_ fields */
#define ZLIST_NAMED_DECL(objtypedef_, field_) \
    zlist ## objtypedef_ ## field_ ## _t


/*
Get the next/previous element in the list. Returns NULL if at end/beginning.
*/
#define ZLIST_NAMED_NEXT(objptr_, field_) ZLINK_NAMED_NEXT(objptr_, field_)
#define ZLIST_NAMED_PREV(objptr_, field_) ZLINK_NAMED_PREV(objptr_, field_)

/*
Get the first/last element of a list. Returns null pointer if empty list.
*/
#define ZLIST_NAMED_HEAD(listptr_, field_)	((listptr_)->zlist_head)
#define ZLIST_NAMED_TAIL(listptr_, field_)	((listptr_)->zlist_tail)

/* 
Init a list as an empty list
*/
#define ZLIST_NAMED_INIT(listptr_) \
    do { \
	(listptr_)->zlist_head = 0; \
	(listptr_)->zlist_tail = 0; \
	(listptr_)->zlist_count = 0; \
    } while (0)
 
/* 
Return number of items in the list 
The 0 is added to make it a rvalue.
*/
#define ZLIST_NAMED_COUNT(listptr_, field_)	(0 + (listptr_)->zlist_count)

/* 
Return 1 if list is empty. 
*/
#define ZLIST_NAMED_ISEMPTY(listptr_, field_) \
    (ZLIST_NAMED_COUNT(listptr_,field_) == 0)


/*
Removes all elements from a list. Directly accesses fields to make it a little
faster than explicitly looping and unlinking. func_ is a macro or function
that is called with a ptr to each element after it is unlinked from the list.
It may do what it wants with the element (eg. deallocate it).
*/
#define ZLIST_NAMED_DESTROY(listptr_, objtypedef_, field_, func_) \
    do { \
	register objtypedef_ *p_; \
	register objtypedef_ *q_; \
	p_ = (listptr_)->zlist_head; \
	while (p_) { \
	    q_ = p_->field_.zl_next; \
	    ZLINK_NAMED_INIT(p_, field_); \
	    func_(p_); \
	    p_ = q_; \
	} \
	(listptr_)->zlist_head = (listptr_)->zlist_tail = 0; \
	(listptr_)->zlist_count = 0; \
    } while (0)

/*
Insert an object after a given object. If curobjptr_ is the last object in
the list, the new object becomes the new tail of the list.
*/
#define ZLIST_NAMED_INSERTAFTER(listptr_, curobjptr_, newobjptr_, field_) \
    do { \
	ZLINK_NAMED_POSTATTACH(curobjptr_, newobjptr_, field_); \
	if ((listptr_)->zlist_tail == (curobjptr_)) { \
	    (listptr_)->zlist_tail = (newobjptr_); \
	} \
	(listptr_)->zlist_count++; \
    } while (0)

/*
Insert an object before a given object. If curobjptr_ is the head of
the list, the new object becomes the new head of the list.
*/
#define ZLIST_NAMED_INSERTBEFORE(listptr_, curobjptr_, newobjptr_, field_) \
    do { \
	ZLINK_NAMED_PREATTACH(curobjptr_, newobjptr_, field_) \
	if ((listptr_)->zlist_head == (curobjptr_)) { \
	    (listptr_)->zlist_head = (newobjptr_); \
	} \
	(listptr_)->zlist_count++; \
    } while (0)


#define ZLIST_NAMED_REMOVE(listptr_, objptr_, field_) \
    do { \
	if ((listptr_)->zlist_head == (objptr_)) \
	    (listptr_)->zlist_head = ZLINK_NAMED_NEXT(objptr_, field_); \
	if ((listptr_)->zlist_tail == (objptr_))  \
	    (listptr_)->zlist_tail = ZLINK_NAMED_PREV(objptr_,field_); \
	ZLINK_NAMED_UNLINK(objptr_, field_); \
	(listptr_)->zlist_count--; \
    } while (0)


/* Add an element to the tail of a list. */
#define ZLIST_NAMED_APPEND(listptr_, objptr_, field_) \
    do { \
	if (ZLIST_NAMED_ISEMPTY((listptr_), field_)) { \
	    (listptr_)->zlist_tail = (objptr_); \
	    (listptr_)->zlist_head = (objptr_); \
	    ZLINK_NAMED_INIT((objptr_), field_); \
	} \
	else { \
	    ZLINK_NAMED_POSTATTACH((listptr_)->zlist_tail, objptr_, field_); \
	    (listptr_)->zlist_tail = (objptr_); \
	} \
	(listptr_)->zlist_count++; \
    } while (0)


/*
Add an element to the head of a list
*/
#define ZLIST_NAMED_PREPEND(listptr_, objptr_, field_) \
    do { \
	if (ZLIST_NAMED_ISEMPTY((listptr_), field_)) { \
	    (listptr_)->zlist_tail = (objptr_); \
	    (listptr_)->zlist_head = (objptr_); \
	    ZLINK_NAMED_INIT((objptr_), field_); \
	} \
	else { \
	    ZLINK_NAMED_PREATTACH((listptr_)->zlist_head, objptr_, field_); \
	    (listptr_)->zlist_head = (objptr_); \
	} \
	(listptr_)->zlist_count++; \
    } while (0)

/*
Move an existing element of the list to the head of the list.
TBD - rewrite to make more efficient by directly manipulating links.
*/
#define ZLIST_NAMED_MOVETOHEAD(listptr_, objptr_, field_) \
    do { \
	ZLIST_NAMED_REMOVE(listptr_, objptr_, field_); \
	ZLIST_NAMED_PREPEND(listptr_, objptr_, field_); \
    } while (0)

/*
Concat list2 to the end of list1
*/
#define ZLIST_NAMED_CONCAT(listptr1_, listptr2_, field_)  \
    do { \
	if (! ZLIST_NAMED_ISEMPTY(listptr2_, field_)) { \
	    if (! ZLIST_NAMED_ISEMPTY(listptr1_, field_)) { \
		(listptr1_)->zlist_tail->field_.zl_next = (listptr2_)->zlist_head; \
		(listptr2_)->zlist_head->field_.zl_prev = (listptr1_)->zlist_tail; \
	    } \
	    else { \
		(listptr1_)->zlist_head = (listptr2_)->zlist_head; \
	    } \
	    (listptr1_)->zlist_tail = (listptr2_)->zlist_tail; \
	    (listptr1_)->zlist_count += (listptr2_)->zlist_count; \
	    (listptr2_)->zlist_tail = 0; \
	    (listptr2_)->zlist_head = 0; \
	    (listptr2_)->zlist_count = 0; \
	} \
    } while (0)


/*
Searches down the list, starting at objptr_ executing cmpmacro_ which must
be a macro or function takes two parameters - the first is the list element
being examined, the second is cmpmacroarg_ which is passed through unchanged.
The macro should return 0 if match succeeds, else non-0.
objptr_ will be set to the first element for which the match succeeded or
to null if non match occurred.
*/
#define ZLIST_NAMED_FIND(objptr_, field_, cmpmacro_, cmpmacroarg_) \
    ZLINK_NAMED_FIND(objptr_, field_, cmpmacro_, cmpmacroarg_)

#define ZLIST_NAMED_SORT_DECL(sortfunc_, objtypedef_, field_) \
    void sortfunc_ (ZLIST_NAMED_DECL(objtypedef_, field_) *listPtr, \
		int (APNCALLBACK *cmpfunc_)(objtypedef_ *, objtypedef_ *))

#define ZLIST_NAMED_SORT_DEFINE(sortfunc_, objtypedef_, field_) \
    ZLIST_NAMED_SORT_DECL(sortfunc_, objtypedef_, field_) \
{ \
    /* This is based on a public domain quicksort routine to be used to */ \
    /* sort by Jon Guthrie. */ \
    objtypedef_ *current; \
    objtypedef_ *head; \
    objtypedef_ *pivot; \
    objtypedef_ *temp; \
    int     result; \
    ZLIST_NAMED_DECL(objtypedef_, field_) low_list, high_list; \
    int i; \
    int count; \
 \
    count = ZLIST_NAMED_COUNT(listPtr, field_); \
    if (count <= 1) \
	return; \
 \
    /* Find the first element that doesn't have the same value as the first */ \
    head = ZLIST_NAMED_HEAD(listPtr, field_); \
    current = head; \
    i = count; \
    do { \
	if (--i == 0) \
	    return; \
	ZLIST_ASSERT(current); \
	current = ZLIST_NAMED_NEXT(current, field_); \
    }   while(0 == (result = cmpfunc_(head, current))); \
 \
    /*  pivot value is the lower of the two.  This insures that the sort */ \
    /*  will always terminate by guaranteeing that there will be at least*/ \
    /*  one member of both of the sublists. */ \
    if(result > 0) \
        pivot = current; \
    else \
        pivot = head; \
 \
    /* Initialize the sublist pointers */ \
    ZLIST_NAMED_INIT(&low_list); \
    ZLIST_NAMED_INIT(&high_list); \
 \
    /* Now, separate the items into the two sublists */ \
    current = head; \
    i = count; \
    while (i--) { \
	ZLIST_ASSERT(current); \
	temp = ZLIST_NAMED_NEXT(current, field_); \
	ZLIST_NAMED_REMOVE(listPtr, current, field_); \
	if(cmpfunc_(pivot, current) < 0) { \
	    /* add one to the high list */ \
	    ZLIST_NAMED_PREPEND(&high_list, current, field_); \
	} \
	else { \
	    /* add one to the low list */ \
	    ZLIST_NAMED_PREPEND(&low_list, current, field_); \
	} \
	current = temp; \
    } \
 \
    /* And, recursively call the sort for each of the two sublists. */ \
    sortfunc_(&low_list, cmpfunc_); \
    sortfunc_(&high_list, cmpfunc_); \
 \
    /* put the "high" list after the end of the "low" list. */ \
    ZLIST_ASSERT(ZLIST_NAMED_ISEMPTY(listPtr, field_)); \
    ZLIST_NAMED_CONCAT(listPtr, &low_list, objtypedef_, field_); \
    ZLIST_NAMED_CONCAT(listPtr, &high_list, objtypedef_, field_); \
    return; \
}




/*
Define a list with default field names.
*/
#define ZLINKNAME__	zlnk__
/* Can't use ZLINKNAME__ in the next def because of the order that the
  CPP ## operator and substitution are done.
*/
#define ZLIST_DECL(objtype_)	ZLIST_NAMED_DECL(objtype_, zlnk__)

#define ZLINK_DECL(objtype_)	ZLINK_NAMED_DECL(objtype_, ZLINKNAME__)

#define ZLINK_INIT(objptr_)	ZLINK_NAMED_INIT((objptr_), ZLINKNAME__)

#define ZLINK_NEXT(objptr_) 	ZLINK_NAMED_NEXT((objptr_), ZLINKNAME__)

#define ZLINK_PREV(objptr_) 	ZLINK_NAMED_PREV((objptr_), ZLINKNAME__)

#define	ZLINK_PREATTACH(curobjptr_, newobjptr_) \
    ZLINK_NAMED_PREATTACH(curobjptr_, newobjptr_, ZLINKNAME__)

#define	ZLINK_POSTATTACH(curobjptr_, newobjptr_) \
    ZLINK_NAMED_POSTATTACH(curobjptr_, newobjptr_, ZLINKNAME__)

#define ZLINK_UNLINK(objptr_)	ZLINK_NAMED_UNLINK((objptr_), ZLINKNAME__)

#define ZLINK_MOVETOFRONT(headptr_, objptr_) \
    ZLINK_NAMED_MOVETOFRONT((headptr_), (objptr_), ZLINKNAME__)

#define ZLINK_SPLIT(objptr_)	ZLINK_NAMED_SPLIT((objptr_), ZLINKNAME__)

#define ZLINK_FIRST(objptr_)	ZLINK_NAMED_FIRST(objptr_, ZLINKNAME__)

#define ZLINK_LAST(objptr_)	ZLINK_NAMED_LAST(objptr_, ZLINKNAME__)

#define ZLINK_FIND(objptr_, cmpmacro_, macroarg_) \
    ZLINK_NAMED_FIND(objptr_, ZLINKNAME__, cmpmacro_, macroarg_)

#define ZLINK_SORT_DECL(sortfunc_, objtypedef_) \
    ZLINK_NAMED_SORT_DECL(sortfunc_, objtypedef_, ZLINKNAME__)

#define ZLINK_SORT_DEFINE(sortfunc_, objtypedef_) \
    ZLINK_NAMED_SORT_DEFINE(sortfunc_, objtypedef_, ZLINKNAME__)

#define ZLIST_CREATE_TYPEDEFS(objtypedef_) \
    ZLIST_NAMED_CREATE_TYPEDEFS(objtypedef_, ZLINKNAME__)

#define ZLIST_NEXT(objptr_) 	ZLIST_NAMED_NEXT((objptr_), ZLINKNAME__)

#define ZLIST_PREV(objptr_) 	ZLIST_NAMED_PREV((objptr_), ZLINKNAME__)

#define ZLIST_HEAD(listptr_) 	ZLIST_NAMED_HEAD((listptr_), ZLINKNAME__)

#define ZLIST_TAIL(listptr_) 	ZLIST_NAMED_TAIL((listptr_), ZLINKNAME__)

#define ZLIST_INIT(listptr_)     ZLIST_NAMED_INIT(listptr_)

#define ZLIST_COUNT(listptr_) 	ZLIST_NAMED_COUNT((listptr_), ZLINKNAME__)

#define ZLIST_ISEMPTY(listptr_)	ZLIST_NAMED_ISEMPTY((listptr_), ZLINKNAME__)

#define ZLIST_DESTROY(listptr_, objtypedef_, func_) \
	ZLIST_NAMED_DESTROY(listptr_, objtypedef_, ZLINKNAME__, func_)

#define ZLIST_INSERTAFTER(listptr_, curobjptr_, newobjptr_) \
    ZLIST_NAMED_INSERTAFTER((listptr_), (curobjptr_), (newobjptr_), ZLINKNAME__)

#define ZLIST_INSERTBEFORE(listptr_, curobjptr_, newobjptr_) \
    ZLIST_NAMED_INSERTBEFORE((listptr_), (curobjptr_), (newobjptr_), ZLINKNAME__)

#define ZLIST_REMOVE(listptr_, objptr_) \
    ZLIST_NAMED_REMOVE((listptr_), (objptr_), ZLINKNAME__)

#define ZLIST_APPEND(listptr_, objptr_) \
    ZLIST_NAMED_APPEND((listptr_), (objptr_), ZLINKNAME__)

#define ZLIST_PREPEND(listptr_, objptr_) \
    ZLIST_NAMED_PREPEND((listptr_), (objptr_), ZLINKNAME__)

#define ZLIST_MOVETOHEAD(listptr_, objptr_) \
    ZLIST_NAMED_MOVETOHEAD((listptr_), objptr_, ZLINKNAME__)

#define ZLIST_CONCAT(listptr1_, listptr2_) \
    ZLIST_NAMED_CONCAT(listptr1_, listptr2_, ZLINKNAME__)

#define ZLIST_FIND(objptr_, cmpmacro_, cmpmacroarg_) \
    ZLIST_NAMED_FIND((objptr_), ZLINKNAME__, cmpmacro_, (cmpmacroarg_))

#define ZLIST_SORT_DECL(sortfunc_, objtypedef_) \
    ZLIST_NAMED_SORT_DECL(sortfunc_, objtypedef_, ZLINKNAME__)

#define ZLIST_SORT_DEFINE(sortfunc_, objtypedef_) \
    ZLIST_NAMED_SORT_DEFINE(sortfunc_, objtypedef_, ZLINKNAME__)


#endif /* ZLIST_H */
