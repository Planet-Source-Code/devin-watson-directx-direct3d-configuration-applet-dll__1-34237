D3DConfig -- A multi-purpose DirectX/Direct3D configuration applet for Windows.

I created D3DConfig as a wrapper around many of the more mundane and often complex functions for enumerating devices and retrieving their individual properties to set up DirectX. Also, I wanted to create an interface which was portable between applications, one that any number of DirectX-based apps could use to let the user set up his/her machine for use of the application.

This is the same as many game configuration screens you may have already seen, such as in Half-Life and Freespace 2 to name a couple.

USAGE
-----

Currently, there is only 1 class within the DLL, D3DConfigForm. This class represents the actual window, and also includes some events which you may want to trap for. In order to use all of the features properly, make sure you use the WithEvents declarator with the Dim statement for D3DConfigForm.

PROPERTIES
----------

Caption -- Sets/gets the caption in the title bar of the applet window. By default, it is "Direct3D Configuration"

Visible -- Sets/gets the visible state of the applet window.

METHODS
-------

Show() -- Makes the applet window visible, and also brings it to the foreground.
Refresh() -- Refreshes the list of adapters.

EVENTS
------

OnCloseClick() -- This event is fired when the user clicks on the "Close" button.
OnOkClick() -- This event is fired when the user clicks okay. Returns which adapter they selected, which resolution mode for that adapter, and whether or not they want to run in Fullscreen mode.