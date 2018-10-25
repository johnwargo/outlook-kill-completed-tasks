# Outlook Kill Completed Tasks

My wife's computer died, and as I transferred her data to her new system, I noticed that Microsoft Outlook reported that she had 348,000 completed tasks in her tasks folder. Now my wife isn't a prolific task completer, so I knew something else was wrong here. As I looked through her task list, I noticed that many of them were duplicates as she only manages something like 10 tasks. Sigh.

Well, I tried to select all, then delete them, but Outlook couldn't handle that. Next I tried deleting them in batches, but since several of the actual tasks were recurring tasks, that wouldn't work either as I would have to click the Outlook confirmation dialog's **Yes** button 348,000 times to delete them.

Fortunately, I knew how to maniuplate Outlook items programatically, so this simple utility app was born. The code finds the tasks folder in Outlook, then deletes all completed tasks. I imagine it will run for days on my wife's new system.

I didn't build in a way to cancel processing, so once it starts, it will run until it completes or until you kill the task.

The source code, as is, doesn't actually delete anything. I've commented out the line that does it, it looks like this:

``` Pascal
// Uncomment the following line when you're ready to delete entries
// oiItem.Delete;
```

With that commented out, you can run the program on your system and see what it would do if it ran for real. Once you're certain you're ready to delete (there is no going back) then uncomment the `oiItem.Delete` line line like so:

``` Pascal
// Uncomment the following line when you're ready to delete entries
oiItem.Delete;
```

Build and execute the program to get to work.

---

By [John M. Wargo](http://www.johnwargo.com) - If you find this code useful, and feel like thanking me for providing it, please consider making a purchase from [my Amazon Wish List](https://amzn.com/w/1WI6AAUKPT5P9). You can find information on many different topics on my [personal blog](http://www.johnwargo.com). Learn about all of my publications at [John Wargo Books](http://www.johnwargobooks.com).