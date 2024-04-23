# AIChatInExcel

## 这什么东西

就是在Excel里用 宏 调用AI接口。

这样很方便的就能拿到 AI 返回的内容。

而且 很容易篡改AI的回复，这样AI的某些拒绝就被我们认为的修改成同意的内容。突破了AI的角色限制。

同时能够获得多个回复，取最好的内容，拼接到一起，有助于 AI 后续的回复 符合我们的期望。

能够在AI跑偏之后，仅通过文本操作，就能让AI回到 正确的问答路上。不像其他UI，很难纠正AI。



## 怎么使用？

1、用WPS打开。在 安全警告 中点击 启用宏。（不信任的话，你可以先看宏的代码，然后再启用）

2、从KIMI 申请到 apiKey 填到 工作表'Config_KIMI' 的 apiKey 的 value项里。

3、切换到 工作表'AIAgentTemplate'，直接点击 send 。保证最后一行是 user 且 content 里 要有内容就行。

4、等差不多1分钟，如果文本长的话，需要等久一点，只要出现 console的提示，就是发出去了。






## What is this?

It's a macro in Excel that calls an AI interface.

This is very convenient as you can easily get the content returned by the AI.

Moreover, it's very easy to tamper with AI's reply, so that some of AI's refusal can be modified into the content we agree with, breaking through the role restrictions of AI.

At the same time, you can obtain multiple replies, take the best content, and splice them together, which is conducive to the subsequent replies of AI being in line with our expectations.

You can get AI back on the right track of Q&A through text operation alone after it goes astray, which is not easy with other UIs.

## How to use it?

1. Open with WPS. Click "Enable macro" in the security warning. (If you don't trust it, you can first look at the macro code and then enable it)

2. Apply for an apiKey from KIMI and fill it into the value item of apiKey in the Config_KIMI worksheet.

3. Switch to the AIAgentTemplate worksheet and click send. Just make sure the last line is user and there is content in the content.

4. Wait for about 1 minute. If the text is long, it will take longer. As long as there is a console prompt, it means it has been sent out.


for lazy me,hope you good luck.


![7ac33b3a67d20f366f16d9440abe975](https://github.com/lautherf/AIChatInExcel/assets/6778243/60d6dea0-a8fe-4e6f-8d03-edb8db1bbc9a)

