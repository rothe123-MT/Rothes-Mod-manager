Download the 3 files the unrealpak.exe is redistributle from Unreal Engine then run the Run_Mod_List_With_Report.bat file to check out the program.

Playlist Original Load Order Feature ✅
New Functionality:
Playlist Creation: When creating a new playlist, it now stores the original load orders for each mod in originalLoadOrders property
Playlist Overwrite: When overwriting an existing playlist, it updates the original load orders
Playlist Editor: When editing a playlist, it captures the current load orders as the new original
Playlist Apply: When applying a playlist, it restores the original load orders that were stored when that playlist was created
How It Works:
When you create a playlist:

Captures current load orders of all enabled mods
Stores them in playlist.originalLoadOrders[FolderKey]
Preserves these values as the "original" state for that playlist
When you apply a playlist:

Uses stored original load orders instead of current mod load orders
Restores each mod to the exact load order it had when that playlist was created
Maintains playlist independence - each playlist has its own original state
When you edit a playlist:

Updates the original load orders to match the current state
Future applications will use the newly saved original values
Benefits:
Playlist Independence: Each playlist maintains its own original load order
No Cross-Contamination: Changing load orders in one playlist doesn't affect others
Reliable Restoration: Always restores to the exact state when playlist was created
Backward Compatible: Existing playlists will work (they'll just use current load orders as fallback)
Now when you switch between playlists, each one will restore to its original load order that was set when that specific playlist was created!
previous above

2.4
Today's Work Summary 📋
🎯 Main Objective Achieved:
Successfully implemented Mass Enable Change functionality for bulk mod management with proper UI positioning and full functionality.

✅ Major Accomplishments:
1. Core Bug Fixes:
Fixed ContainsKey error - resolved originalLoadOrders PSCustomObject vs Hashtable issue
Fixed playlist apply - enabled status now applies automatically to correct column
Fixed column indexing - ensured all references point to correct columns
2. Mass Enable Change Feature:
Added checkbox column at end of grid (column 6) to avoid index conflicts
Implemented toggle functionality:
Disabled mods → Enabled when checked
Enabled mods → Disabled when checked
Revert to original when unchecked
Updated column header to "Mass Enable Change" for clarity
3. UI Improvements:
Repositioned logo outside grid area to the right with auto-sizing
Fixed Reset Playlists button:
Removed duplicate button that was causing conflicts
Updated confirmation message to "Are you sure? This will remove all playlists."
Positioned button at (940, 605) near Delete Selected button
Fixed overlap issues with proper spacing
4. Column Structure:
0: Name | 1: Load Order | 2: Original Load Order | 3: Status | 4: Enabled | 5: Details | 6: Mass Enable Change
🔧 Technical Details:
CellValueChanged event handler for real-time checkbox functionality
Proper column index management to prevent conflicts
Original state preservation for revert functionality
Error handling for playlist operations
📝 What's Ready:
✅ All playlist functionality working correctly
✅ Mass Enable Change toggle fully functional
✅ UI properly aligned with no overlaps
✅ Confirmation dialogs working properly
✅ Logo positioned outside grid area

The script is now ready with enhanced bulk mod management capabilities!

2.5

Summary of Today's Work 📋
✅ All Issues Resolved:
Enabled Column Sortable - Added SortMode property
Report Generation Fixed - Only creates when Update Report button clicked
Grid Sort Order Fixed - Populates with unsorted mods, relies on DataGridView sorting
Grid Refresh Error Fixed - Added null check before calling Refresh()
Report Order Consistency - Removed LoadOrder sort, preserves grid display order
🎯 Current State:
Grid: Shows mods in original order, respects user column sorting
Report: Generates in exact same order as grid display
Mass Enable Change: Full toggle functionality working
UI: All buttons properly positioned and functional
The mod manager now has consistent sorting between grid and report, with full user control over display order!
