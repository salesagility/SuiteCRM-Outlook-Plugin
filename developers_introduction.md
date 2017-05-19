# Introduction for Developers

## History

The SuiteCRM Outlook Add-in has a history, and it's fairly important to understand that when working on it. The earliest commits in the current repository are by [Will Rennie](https://github.com/willrennie), but he wasn't responsible for the original design, instead inheriting a pre-existing codebase. Will improved the codebase to some extent between May 2014 and January 2017. [Andrew Forrest](https://github.com/bunsen32) then took over briefly. Andrew is a very experienced and expert software engineer who actually likes C#, and did a considerable amount of refactoring which improved the code greatly. In March 2017 [Simon Brooke](https://github.com/simon-brooke) took over. Simon is an aging and grumpy Lisp hacker who loathes everything produced by Microsoft generally, and C# particularly, with a passion. Simon has hacked around the codebase and made stuff work.

But this tangled history leads to a variety of styles, and this can be hard to follow.

Parts of the original codebase were startlingly bad, with tangled nests of GOTO statements; Simon has a suspicion that the original code was produced using a nasty point-and-drool paint-an-app tool by some *Microsoft Certified Software Professional*, probably originally as Visual BASIC, and then automatically converted to C# with some equally nasty tool. He can find no other explanation for its deep horridness. There was no design documentation, or if there was it has long been lost. There was no inline documentation. Consequently everyone working on it since has been working blind; we've all tried to improve it in our different ways, and on the whole the general quality is now much better.

Parts of the ugliness remain, however.

The original code makes heavy use of Hungarian notation, and nothing conformed to normal C# naming conventions. The conventional Microsoft Settings class was not used, but instead a custom 'clsSettings'; and classes were originally primarily in the SuiteCRMAddIn namespace with no real structure.

This is being improved, gradually. Newly created forms and dialogs, for example, are now in the namespace SuiteCRMAddin.Dialogs, and have class names ending in **Dialog**, whereas older forms and dialogs are found in SuiteCRMAddIn with names prefixed **frm**. The intention is that they should gradually be migrated.

## Documentation

Documentation is partial and incomplete; and, since we don't know the original designers intention and have had to infer it from the code, some may be mistaken or just plain wrong. We are gradually trying to improve this situation.

Broadly, all inline documentation has been written by Simon; at the point that he took over there was very little inline documentation at all. He is to be blamed, therefore, for all the bits that are wrong or out of date.

HTML documentation is generated from the inline documentation using [doxygen](https://www.stack.nl/~dimitri/doxygen/). **The HTML documentation should never be directly edited**. Instead please edit inline documentation in the code, and regenerate with doxygen. A convention has been adopted of adding a file *Documentation.cs* to each package, to contain package level documentation. By convention this file should contain no class or other code.
