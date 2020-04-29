# Chapter 19
### Guidelines
- DO NOT fall into the common error of believing that more threads always make code faster.
- DO carefully measure performance when attempting to speed up processor-bound problems through multithreading.
### Guidelines
- DO NOT make an unwarranted assumption that any operation that is seemingly atomic in single threaded code will be atomic in multithreaded code.
- DO NOT assume that all threads will observe all side effects of operations on shared memory in a consistent order.
- DO ensure that code that concurrently acquires multiple locks always acquires them in the same order.
- AVOID all race conditions—that is, conditions where program behavior depends on how the operating system chooses to schedule threads.
### Guidelines
- AVOID writing programs that produce unhandled exceptions on any thread.
- CONSIDER registering an unhandled exception event handler for debugging, logging, and emergency shutdown purposes.
- DO cancel unfinished tasks rather than allowing them to run during application shutdown.
### Guidelines
- DO cancel unfinished tasks rather than allowing them to run during application shutdown.
### Guidelines
- DO inform the task factory that a newly created task is likely to be long-running so that it can manage it appropriately.
- DO use TaskCreationOptions.LongRunning sparingly.
### Guidelines
- AVOID calling Thread.Sleep() in production code.
- DO use tasks and related APIs in favor of System.Theading classes such as Thread and ThreadPool.
