# Differences between @stoneage7's upstream repository and this fork

- Support for TreeSet, which like Java's TreeSet, internally uses an instance of TreeMap with the value being a marker value (Empty in this case).
- remove(key) on TreeMap now returns the node removed, or Nothing if there was no candidate to be removed, as opposed to failing with an error if the key does not exist.
- A few methods and classes have been renamed.
- Dumping takes a path, as opposed to using a built in path relative to the workbook.
- Support for custom comparators by setting the key_comparator (TreeMap) or entry_comparator (TreeSet) property, as opposed to having to change TreeMap.cmp_key().
- VB6 compatible (no longer uses VBA specific functions).

# Motivation (as per @stoneage7's upstream repository)

This class was written to work around limitations of the standard Collection class:

- find() returns a reference to existing TreeMapNode -> its payload can be changed.
- can do partial iteration (For Each with Collection object can only do full; iteration with Collection.Item is O(N^2)).
- duplicate keys don't have to be treated as errors.
- ~~add() can use simple flat array as key (as long as array members can be compared using '<' and '=').~~ _(This is specific to the default IKeyComparator implementation)_
- ~~any object can be key as long as TreeMap.cmp_key() is adapted.~~ _(Updated to use KeyComparator interface)_

# Classes and Methods

## Class TreeMap
    - add(key As Variant, value As Variant) As TreeMapNode
    - find(key As Variant) As TreeMapNode
    - count() As Long
    - Get duplication_mode() As TreeMapDuplicationMode
    - Let duplication_mode(a_duplicaiton_mode As TreeMapDuplicationMode)
    - Get key_comparator() As IVariantComparator
    - Set key_comparator(a_key_comparator As IVariantComparator)
    - remove(key As Variant) As TreeMapNode
    - create_in_order_cursor(Optional from_key As Variant) As TreeMapInOrderCursor
    - dump(path AS String, Optional N As TreeMapNode)

## Class TreeMapNode

All members are public for simplicity.

    - payload as Variant

## Class TreeMapInOrderCursor
    - next_node() As TreeMapNode
    - prev_node() As TreeMapNode
    - first_node() As TreeMapNode
    - last_node() As TreeMapNode
    - start(start_at As TreeMapNode)

All return Nothing when there is no such node to move to.

## Class TreeSet
    - add(entry As Variant) As Boolean (returns true if the set did not already contain the entry, false otherwise.)
    - find(entry As Variant) As Boolean (returns true if the set contains a matching entry, false otherwise.)
    - count() As Long
    - Get entry_comparator() As IVariantComparator
    - Set entry_comparator(a_entry_comparator As IVariantComparator)
    - remove(entry As Variant) As Boolean (returns true if the set contained a matching entry that was removed, false otherwise).
    - create_in_order_cursor(Optional from_entry As Variant) As TreeSetInOrderCursor
    - dump(path AS String, Optional N As Variant)

## Class TreeSetInOrderCursor
    - next_entry() As Variant
    - prev_entry() As Variant
    - first_entry() As Variant
    - last_entry() As Variant
    - start_using(map_cursor_ As TreeMapInOrderCursor)

All return Nothing when there is no such node to move to.

## Interface IVariantComparator
    - compare(v1 As Variant, v2 As Variant) As Long

## Class SimpleVariantComparator implements IVariantComparator
    - IVariantComparator_compare(v1 As Variant, v2 As Variant) As Long (The existing cmp_key method from the upstream TreeMap class, extracted into a new class)

## Enum TreeMapDuplicationMode
    - TreeMapDuplicationMode_Ignore
    - TreeMapDuplicationMode_RaiseError
    - TreeMapDuplicationMode_Overwrite