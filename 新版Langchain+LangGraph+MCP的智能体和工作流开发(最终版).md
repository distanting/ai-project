**新版Langchain + LangGraph +MCP 的智能体和工作流开发**

**讲师：肖斌， 今晚VIP直播课，8点05分开始**

**内容：**

**1、第二期直播的内容预告**

**2、学习路线的要求**

**3、上期的部分同学简历和工作情况汇报**

**4、多（全）模态大模型部署 + 应用开发**

**第一章、大模型选择和私有化部署**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/1.png)

**强烈建议：在大模型开发中，要熟悉各种顶级的AI大模型的使用和微调。包括：gpt-4o， gpt-o1-mini， gpt-o3，claude-3.5-sonnet，claude-3.7-sonnet，gemini-1.5, deepseek系列, qwen系列。**

**1、Deepseek-V3**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/2.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/3.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/4.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/5.png)

**2、Qwen3**

4月29日，发布了 Qwen3 系列模型。我们的旗舰模型\*\*Qwen3-235B-A22B\*\*在编码、数学、通用能力等基准测试中，与 DeepSeek-R1、o1、o3-mini、Grok-3 和 Gemini-2.5-Pro 等其他顶级模型相比，取得了极具竞争力的成绩。此外，小型 MoE 模型\*\*Qwen3-30B-A3B 的\*\*激活参数量是 QwQ-32B 的 10 倍，即使是像 Qwen3-4B 这样的微型模型，其性能也能与 Qwen2.5-72B-Instruct 相媲美。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/6.png)

**开源的Qwen3-235B-A22B**是一个大型模型，总参数量达 2350 亿，激活参数量达 220 亿；Qwen3-30B-A3B是一个较小的 MoE 模型，总参数量达 300 亿，激活参数量达 30 亿。此外，六个Dense模型也已开源，包括 **Qwen3-32B** 、 **Qwen3-14B** 、 **Qwen3-8B** 、 **Qwen3-4B** 、Qwen3-1.7B 和 **Qwen3-0.6B** ，它们均遵循 Apache 2.0 许可证。

**Qwen3 的亮点包括：**

* **各种尺寸的密集和混合专家 (MoE) 模型** ，有 0.6B、1.7B、4B、8B、14B、32B 和 30B-A3B、235B-A22B 可供选择。
* **思维模式**（用于复杂的逻辑推理、数学和编码）和 **非思维模式** （用于高效、通用的聊天）之间的无缝切换，确保在各种场景下实现最佳性能。
* **推理能力大幅增强** ，在数学、代码生成、常识逻辑推理等方面超越了之前的QwQ（思维模式）和Qwen2.5指令模型（非思维模式）。
* **卓越的人类偏好一致性** ，擅长创意写作、角色扮演、多轮对话和指令遵循，提供更自然、更具吸引力和身临其境的对话体验。
* **精通Agent能力** ，能够以思考和非思考两种模式与外部工具精准集成，并增强了对 MCP 的支持，在基于Agent的复杂任务中取得开源模型的领先性能。
* **支持 100 多种语言和方言，支持**119 种语言和方言

**3、私有化部署**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/7.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/8.png)

**显存计算器**

* 大模型显存需求分析工具 | LlamaFactory | LlamaFactory:

**https://www.llamafactory.cn/tools/gpu-memory-estimation.html**

**一、下载模型**

打开网站：https://www.modelscope.cn/organization/Qwen?tab=collection

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/9.png)

|  |
| --- |
| Shell from modelscope import snapshot\_download  # model\_dir = snapshot\_download('ZhipuAI/glm-4-9b-chat', cache\_dir='/root/autodl-tmp/models', revision='master') # model\_dir = snapshot\_download('deepseek-ai/DeepSeek-R1-Distill-Qwen-7B', cache\_dir='/root/autodl-tmp/models', revision='master') # model\_dir = snapshot\_download('LLM-Research/Meta-Llama-3.1-8B-Instruct', cache\_dir='/root/autodl-tmp/models', revision='master') # model\_dir = snapshot\_download('Qwen/Qwen2.5-7B-Instruct', cache\_dir='/root/autodl-tmp/models', revision='master')  #模型下载 # from modelscope import snapshot\_download # model\_dir = snapshot\_download('LLM-Research/Meta-Llama-3.1-8B-Instruct')  # model\_dir = snapshot\_download('Qwen/Qwen3-8B', cache\_dir='/root/autodl-tmp/models', revision='master')  model\_dir = snapshot\_download('Qwen/Qwen2.5-Omni-3B', cache\_dir='/root/autodl-tmp/models', revision='master') |

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/10.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/11.png)

**二、通过vllm server命令部署**

**企业生产环境中，不要使用Ollama部署**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/12.png)

以下是Ollama和vLLM在部署方面的对比表格：

|  |  |  |
| --- | --- | --- |
| **对比维度** | **Ollama** | **vLLM** |
| **定位** | 轻量化本地部署工具，适合个人开发者和快速原型验证 | 高性能推理框架，面向企业级生产环境和高并发场景 |
| **部署复杂度** | 低，支持一键安装和运行（如 ollama run命令） | 中高，需配置Python环境、API服务接口和分布式集群 |
| **硬件要求** | 支持CPU/GPU，最低配置为16GB内存（运行7B模型） | 强制需要GPU（如NVIDIA Tesla系列），显存要求较高 |
| **并发能力** | 有限，适合单会话或少量并发 | 支持高并发，通过连续批处理（Continuous Batching）优化吞吐量1 |
| **资源占用** | 单机环境下资源占用低，启动快 | 资源占用高但利用率优，支持多机多卡扩展 |
| **延迟表现** | 实时交互场景延迟更低 | 通过批处理平衡延迟与吞吐，适合高吞吐场景 |
| **生态支持** | 丰富的预置模型（如Llama系列），支持跨平台（Windows/macOS/Linux） | 聚焦推理优化，企业级功能丰富（如分布式推理、量化支持） |
| **典型安装命令** | curl -fsSL https://ollama.com/install.sh | sh（Linux）1 | pip install vllm + 配置API服务 |
| **适用场景** | 快速原型验证、本地开发、教育演示、资源受限环境 | 企业级API服务、高并发聊天机器人、多GPU集群任务 |
| **模型格式支持** | GGUF格式（适合量化与跨平台） | HuggingFace Transformers格式（.bin/.safetensors） |
| **长上下文支持** | 默认4K-8K tokens | 支持32K-128K tokens（依赖PagedAttention技术） |
| **量化支持** | 自动支持Q4\_0、Q5\_K等量化格式 | 需外部工具（如bitsandbytes）实现量化 |

**vllm的命令说明**

首先需要安装： pip install vllm

[OpenAI 兼容服务器 — vLLM 文档](https://docs.vllm.com.cn/en/latest/serving/openai_compatible_server.html#openai-compatible-server)

usage: vllm serve [-h] [--host HOST] [--port PORT]

[--uvicorn-log-level {debug,info,warning,error,critical,trace}]

[--disable-uvicorn-access-log] [--allow-credentials]

[--allowed-origins ALLOWED\_ORIGINS]

[--allowed-methods ALLOWED\_METHODS]

[--allowed-headers ALLOWED\_HEADERS] [--api-key API\_KEY]

[--lora-modules LORA\_MODULES [LORA\_MODULES ...]]

[--prompt-adapters PROMPT\_ADAPTERS [PROMPT\_ADAPTERS ...]]

[--chat-template CHAT\_TEMPLATE]

[--chat-template-content-format {auto,string,openai}]

[--response-role RESPONSE\_ROLE] [--ssl-keyfile SSL\_KEYFILE]

[--ssl-certfile SSL\_CERTFILE] [--ssl-ca-certs SSL\_CA\_CERTS]

[--enable-ssl-refresh] [--ssl-cert-reqs SSL\_CERT\_REQS]

[--root-path ROOT\_PATH] [--middleware MIDDLEWARE]

[--return-tokens-as-token-ids]

[--disable-frontend-multiprocessing]

[--enable-request-id-headers] [--enable-auto-tool-choice]

[--tool-call-parser {granite-20b-fc,granite,hermes,internlm,jamba,llama3\_json,mistral,phi4\_mini\_json,pythonic} or name registered in --tool-parser-plugin]

[--tool-parser-plugin TOOL\_PARSER\_PLUGIN] [--model MODEL]

[--task {auto,generate,embedding,embed,classify,score,reward,transcription}]

[--tokenizer TOKENIZER] [--hf-config-path HF\_CONFIG\_PATH]

[--skip-tokenizer-init] [--revision REVISION]

[--code-revision CODE\_REVISION]

[--tokenizer-revision TOKENIZER\_REVISION]

[--tokenizer-mode {auto,slow,mistral,custom}]

[--trust-remote-code]

[--allowed-local-media-path ALLOWED\_LOCAL\_MEDIA\_PATH]

[--download-dir DOWNLOAD\_DIR]

[--load-format {auto,pt,safetensors,npcache,dummy,tensorizer,sharded\_state,gguf,bitsandbytes,mistral,runai\_streamer,fastsafetensors}]

[--config-format {auto,hf,mistral}]

[--dtype {auto,half,float16,bfloat16,float,float32}]

[--kv-cache-dtype {auto,fp8,fp8\_e5m2,fp8\_e4m3}]

[--max-model-len MAX\_MODEL\_LEN]

[--guided-decoding-backend GUIDED\_DECODING\_BACKEND]

[--logits-processor-pattern LOGITS\_PROCESSOR\_PATTERN]

[--model-impl {auto,vllm,transformers}]

[--distributed-executor-backend {ray,mp,uni,external\_launcher}]

[--pipeline-parallel-size PIPELINE\_PARALLEL\_SIZE]

[--tensor-parallel-size TENSOR\_PARALLEL\_SIZE]

[--data-parallel-size DATA\_PARALLEL\_SIZE]

[--enable-expert-parallel]

[--max-parallel-loading-workers MAX\_PARALLEL\_LOADING\_WORKERS]

[--ray-workers-use-nsight] [--block-size {8,16,32,64,128}]

[--enable-prefix-caching | --no-enable-prefix-caching]

[--prefix-caching-hash-algo {builtin,sha256}]

[--disable-sliding-window] [--use-v2-block-manager]

[--num-lookahead-slots NUM\_LOOKAHEAD\_SLOTS] [--seed SEED]

[--swap-space SWAP\_SPACE] [--cpu-offload-gb CPU\_OFFLOAD\_GB]

[--gpu-memory-utilization GPU\_MEMORY\_UTILIZATION]

[--num-gpu-blocks-override NUM\_GPU\_BLOCKS\_OVERRIDE]

[--max-num-batched-tokens MAX\_NUM\_BATCHED\_TOKENS]

[--max-num-partial-prefills MAX\_NUM\_PARTIAL\_PREFILLS]

[--max-long-partial-prefills MAX\_LONG\_PARTIAL\_PREFILLS]

[--long-prefill-token-threshold LONG\_PREFILL\_TOKEN\_THRESHOLD]

[--max-num-seqs MAX\_NUM\_SEQS] [--max-logprobs MAX\_LOGPROBS]

[--disable-log-stats]

[--quantization {aqlm,awq,deepspeedfp,tpu\_int8,fp8,ptpc\_fp8,fbgemm\_fp8,modelopt,nvfp4,marlin,gguf,gptq\_marlin\_24,gptq\_marlin,awq\_marlin,gptq,compressed-tensors,bitsandbytes,qqq,hqq,experts\_int8,neuron\_quant,ipex,quark,moe\_wna16,torchao,None}]

[--rope-scaling ROPE\_SCALING] [--rope-theta ROPE\_THETA]

[--hf-token [HF\_TOKEN]] [--hf-overrides HF\_OVERRIDES]

[--enforce-eager]

[--max-seq-len-to-capture MAX\_SEQ\_LEN\_TO\_CAPTURE]

[--disable-custom-all-reduce]

[--tokenizer-pool-size TOKENIZER\_POOL\_SIZE]

[--tokenizer-pool-type TOKENIZER\_POOL\_TYPE]

[--tokenizer-pool-extra-config TOKENIZER\_POOL\_EXTRA\_CONFIG]

[--limit-mm-per-prompt LIMIT\_MM\_PER\_PROMPT]

[--mm-processor-kwargs MM\_PROCESSOR\_KWARGS]

[--disable-mm-preprocessor-cache] [--enable-lora]

[--enable-lora-bias] [--max-loras MAX\_LORAS]

[--max-lora-rank MAX\_LORA\_RANK]

[--lora-extra-vocab-size LORA\_EXTRA\_VOCAB\_SIZE]

[--lora-dtype {auto,float16,bfloat16}]

[--long-lora-scaling-factors LONG\_LORA\_SCALING\_FACTORS]

[--max-cpu-loras MAX\_CPU\_LORAS] [--fully-sharded-loras]

[--enable-prompt-adapter]

[--max-prompt-adapters MAX\_PROMPT\_ADAPTERS]

[--max-prompt-adapter-token MAX\_PROMPT\_ADAPTER\_TOKEN]

[--device {auto,cuda,neuron,cpu,tpu,xpu,hpu}]

[--num-scheduler-steps NUM\_SCHEDULER\_STEPS]

[--use-tqdm-on-load | --no-use-tqdm-on-load]

[--multi-step-stream-outputs [MULTI\_STEP\_STREAM\_OUTPUTS]]

[--scheduler-delay-factor SCHEDULER\_DELAY\_FACTOR]

[--enable-chunked-prefill [ENABLE\_CHUNKED\_PREFILL]]

[--speculative-config SPECULATIVE\_CONFIG]

[--model-loader-extra-config MODEL\_LOADER\_EXTRA\_CONFIG]

[--ignore-patterns IGNORE\_PATTERNS]

[--preemption-mode PREEMPTION\_MODE]

[--served-model-name SERVED\_MODEL\_NAME [SERVED\_MODEL\_NAME ...]]

[--qlora-adapter-name-or-path QLORA\_ADAPTER\_NAME\_OR\_PATH]

[--show-hidden-metrics-for-version SHOW\_HIDDEN\_METRICS\_FOR\_VERSION]

[--otlp-traces-endpoint OTLP\_TRACES\_ENDPOINT]

[--collect-detailed-traces COLLECT\_DETAILED\_TRACES]

[--disable-async-output-proc]

[--scheduling-policy {fcfs,priority}]

[--scheduler-cls SCHEDULER\_CLS]

[--override-neuron-config OVERRIDE\_NEURON\_CONFIG]

[--override-pooler-config OVERRIDE\_POOLER\_CONFIG]

[--compilation-config COMPILATION\_CONFIG]

[--kv-transfer-config KV\_TRANSFER\_CONFIG]

[--worker-cls WORKER\_CLS]

[--worker-extension-cls WORKER\_EXTENSION\_CLS]

[--generation-config GENERATION\_CONFIG]

[--override-generation-config OVERRIDE\_GENERATION\_CONFIG]

[--enable-sleep-mode] [--calculate-kv-scales]

[--additional-config ADDITIONAL\_CONFIG] [--enable-reasoning]

[--reasoning-parser {deepseek\_r1,granite}]

[--disable-cascade-attn]

[--disable-chunked-mm-input [DISABLE\_CHUNKED\_MM\_INPUT]]

[--disable-log-requests] [--max-log-len MAX\_LOG\_LEN]

[--disable-fastapi-docs] [--enable-prompt-tokens-details]

[--enable-server-load-tracking]

**命名参数详解**

* --host主机名。
* --port端口号。

默认值：8000

* --uvicorn-log-level可能选项：debug, info, warning, error, critical, trace

uvicorn 的日志级别。

默认值：“info”

* --disable-uvicorn-access-log禁用 uvicorn 访问日志。

默认值：False

* --allow-credentials允许凭据。

默认值：False

* --allowed-origins允许的来源。

默认值：[‘\*’]

* --allowed-methods允许的方法。

默认值：[‘\*’]

* --allowed-headers允许的标头。

默认值：[‘\*’]

* --api-key如果提供，服务器将要求在标头中提供此密钥。
* --lora-modules：LoRA 模块配置，格式为 ‘name=path’ 或 JSON 格式。示例（旧格式）：<span class="pre">'name=path'</span> 示例（新格式）：
* --prompt-adapters提示适配器配置，格式为 name=path。可以指定多个适配器。
* --chat-template指定模型的聊天模板文件路径或单行形式的模板。
* --chat-template-content-format可能选项：auto, string, openai

在聊天模板中渲染消息内容的格式。

* “string” 将内容渲染为字符串。示例：
* “openai” 将内容渲染为字典列表，类似于 OpenAI 模式。示例：

默认值：“auto”

* --response-role如果 <span class="pre">request.add\_generation\_prompt=true</span>，则返回的角色名称。

默认值：assistant

* --ssl-keyfileSSL 密钥文件的文件路径。
* --ssl-certfileSSL 证书文件的文件路径。
* --ssl-ca-certsCA 证书文件。
* --enable-ssl-refresh当 SSL 证书文件更改时刷新 SSL 上下文

默认值：False

* --ssl-cert-reqs是否需要客户端证书（请参阅 stdlib ssl 模块）。

默认值：0

* --root-path当应用程序位于基于路径的路由代理之后时，FastAPI root\_path。
* --middleware要应用于应用程序的其他 ASGI 中间件。我们接受多个 --middleware 参数。该值应为导入路径。如果提供了函数，vLLM 将使用 <span class="pre">@app.middleware('http')</span> 将其添加到服务器。如果提供了类，vLLM 将使用 <span class="pre">app.add\_middleware()</span> 将其添加到服务器。

默认值：[]

* --return-tokens-as-token-ids当指定 <span class="pre">--max-logprobs</span> 时，将单个 token 表示为 ‘token\_id:{token\_id}’ 形式的字符串，以便可以识别不可 JSON 编码的 token。

默认值：False

* --disable-frontend-multiprocessing如果指定，将在与模型服务引擎相同的进程中运行 OpenAI 前端服务器。

默认值：False

* --enable-request-id-headers如果指定，API 服务器将在响应中添加 X-Request-Id 标头。注意：在高 QPS 下，这会降低性能。

默认值：False

* --enable-auto-tool-choice为支持的模型启用自动工具选择。使用 <span class="pre">--tool-call-parser</span> 指定要使用的解析器。

默认值：False

* --tool-call-parser根据您使用的模型选择工具调用解析器。这用于将模型生成的工具调用解析为 OpenAI API 格式。<span class="pre">--enable-auto-tool-choice</span> 需要此参数。
* --tool-parser-plugin指定工具解析器插件，用于将模型生成的工具解析为 OpenAI API 格式，在此插件中注册的名称可在 <span class="pre">--tool-call-parser</span> 中使用。

默认值：“”

* --model 要使用的 huggingface 模型的名称或路径。

默认值：“facebook/opt-125m”

* --task可能选项：auto, generate, embedding, embed, classify, score, reward, transcription

模型要执行的任务。即使同一个模型可以用于多个任务，每个 vLLM 实例也仅支持一个任务。当模型仅支持一个任务时，可以使用 <span class="pre">"auto"</span> 选择它；否则，您必须明确指定要使用的任务。

默认值：“auto”

* --tokenizer要使用的 huggingface 分词器的名称或路径。如果未指定，将使用模型名称或路径。
* --hf-config-path要使用的 huggingface 配置的名称或路径。如果未指定，将使用模型名称或路径。
* --skip-tokenizer-init跳过分词器和反分词器的初始化。期望输入中提供有效的 prompt\_token\_ids，并且 prompt 为 None。生成的输出将包含 token ID。

默认值：False

* --revision要使用的特定模型版本。它可以是分支名称、标签名称或提交 ID。如果未指定，将使用默认版本。
* --code-revision用于 Hugging Face Hub 上模型代码的特定修订版本。它可以是分支名称、标签名称或提交 ID。如果未指定，将使用默认版本。
* --tokenizer-revision要使用的 huggingface 分词器的修订版本。它可以是分支名称、标签名称或提交 ID。如果未指定，将使用默认版本。
* --tokenizer-mode可能选项：auto, slow, mistral, custom

分词器模式。

* “auto” 将在可用时使用快速分词器。
* “slow” 将始终使用慢速分词器。
* “mistral” 将始终使用 mistral\_common 分词器。
* “custom” 将使用 –tokenizer 选择预注册的分词器。

默认值：“auto”

* --trust-remote-code信任来自 huggingface 的远程代码。

默认值：False

* --allowed-local-media-path允许 API 请求从服务器文件系统指定的目录读取本地图像或视频。这是一个安全风险。仅应在受信任的环境中启用。
* --download-dir下载和加载权重的目录。
* --load-format可能选项：auto, pt, safetensors, npcache, dummy, tensorizer, sharded\_state, gguf, bitsandbytes, mistral, runai\_streamer, fastsafetensors

要加载的模型权重的格式。

* “auto” 将尝试加载 safetensors 格式的权重，如果 safetensors 格式不可用，则回退到 pytorch bin 格式。
* “pt” 将加载 pytorch bin 格式的权重。
* “safetensors” 将加载 safetensors 格式的权重。
* “npcache” 将加载 pytorch 格式的权重，并存储 numpy 缓存以加速加载。
* “dummy” 将使用随机值初始化权重，主要用于性能分析。
* “tensorizer” 将使用 CoreWeave 的 tensorizer 加载权重。有关更多信息，请参阅“示例”部分中的“张量化 vLLM 模型”脚本。
* “runai\_streamer” 将使用 Run:ai Model Streamer 加载 Safetensors 权重。
* “bitsandbytes” 将使用 bitsandbytes 量化加载权重。
* “sharded\_state” 将从预分片检查点文件加载权重，支持高效加载张量并行模型
* “gguf” 将从 GGUF 格式文件加载权重（详细信息请参阅 [ggml-org/ggml](https://github.com/ggml-org/ggml/blob/master/docs/gguf.md)）。
* “mistral” 将从 Mistral 模型使用的合并 safetensors 文件加载权重。

默认值：“auto”

* --config-format可能选项：auto, hf, mistral

要加载的模型配置的格式。

* “auto” 将尝试加载 hf 格式的配置，如果不可用，则尝试加载 mistral 格式

默认值：“ConfigFormat.AUTO”

* --dtype可能选项：auto, half, float16, bfloat16, float, float32

模型权重和激活的数据类型。

* “auto” 将对 FP32 和 FP16 模型使用 FP16 精度，对 BF16 模型使用 BF16 精度。
* “half” 表示 FP16。推荐用于 AWQ 量化。
* “float16” 与 “half” 相同。
* “bfloat16” 在精度和范围之间取得平衡。
* “float” 是 FP32 精度的简写。
* “float32” 表示 FP32 精度。

默认值：“auto”

* --kv-cache-dtype可能选项：auto, fp8, fp8\_e5m2, fp8\_e4m3

kv 缓存存储的数据类型。如果为 “auto”，将使用模型数据类型。CUDA 11.8+ 支持 fp8 (=fp8\_e4m3) 和 fp8\_e5m2。ROCm (AMD GPU) 支持 fp8 (=fp8\_e4m3)

默认值：“auto”

* --max-model-len模型上下文长度。如果未指定，将自动从模型配置中派生。支持 k/m/g/K/M/G 人类可读格式。示例：- 1k → 1000 - 1K → 1024
* --guided-decoding-backend默认情况下，哪个引擎将用于引导解码（JSON 模式/正则表达式等）。当前支持 [mlc-ai/xgrammar](https://github.com/mlc-ai/xgrammar) 和 [guidance-ai/llguidance.Valid](https://github.com/guidance-ai/llguidance.Valid)。有效的后端值包括 “xgrammar”、“guidance” 和 “auto”。使用 “auto” 时，我们将根据请求内容和后端库当前支持的内容做出有主见的决定，因此行为可能会在每个版本中更改。

默认值：“xgrammar”

* --logits-processor-pattern可选的正则表达式模式，用于指定可以使用 logits\_processors 额外完成参数传递的有效 logits 处理器限定名称。默认为 None，表示不允许任何处理器。
* --model-impl可能选项：auto, vllm, transformers

要使用的模型实现。

* “auto” 将尝试使用 vLLM 实现（如果存在），如果 vLLM 实现不可用，则回退到 Transformers 实现。
* “vllm” 将使用 vLLM 模型实现。
* “transformers” 将使用 Transformers 模型实现。

默认值：“auto”

* --distributed-executor-backend可能选项：ray, mp, uni, external\_launcher

用于分布式模型工作程序的后端，可以是 “ray” 或 “mp”（多进程）。如果 pipeline\_parallel\_size 和 tensor\_parallel\_size 的乘积小于或等于可用 GPU 的数量，“mp” 将用于保持在单个主机上处理。否则，如果安装了 Ray，则默认为 “ray”，否则将失败。请注意，tpu 仅支持 Ray 进行分布式推理。

* --pipeline-parallel-size, -pp流水线并行阶段数。

默认值：1

* --tensor-parallel-size, -tp张量并行副本数。

默认值：1

* --data-parallel-size, -dp数据并行副本数。MoE 层将根据 tensor-parallel-size 和 data-parallel-size 的乘积进行分片。

默认值：1

* --enable-expert-parallel对 MoE 层使用专家并行而不是张量并行。

默认值：False

* --max-parallel-loading-workers在多个批次中顺序加载模型，以避免在使用张量并行和大型模型时出现 RAM OOM。
* --ray-workers-use-nsight如果指定，则使用 nsight 分析 Ray 工作程序。

默认值：False

* --block-size可能选项：8, 16, 32, 64, 128

token 块大小，用于 token 的连续块。在 neuron 设备上将被忽略，并设置为 <span class="pre">--max-model-len</span>。在 CUDA 设备上，仅支持最大为 32 的块大小。在 HPU 设备上，块大小默认为 128。

* --enable-prefix-caching, --no-enable-prefix-caching启用自动前缀缓存。使用 <span class="pre">--no-enable-prefix-caching</span> 显式禁用。
* --prefix-caching-hash-algo可能选项：builtin, sha256

设置前缀缓存的哈希算法。选项包括 ‘builtin’（Python 的内置哈希）或 ‘sha256’（抗冲突但具有一定开销）。

默认值：“builtin”

* --disable-sliding-window禁用滑动窗口，限制为滑动窗口大小。

默认值：False

* --use-v2-block-manager[已弃用] 块管理器 v1 已被删除，SelfAttnBlockSpaceManager（即块管理器 v2）现在是默认设置。将此标志设置为 True 或 False 对 vLLM 行为没有影响。

默认值：True

* --num-lookahead-slots推测解码所需的实验性调度配置。这将在未来被推测配置取代；它目前的存在是为了启用正确性测试。

默认值：0

* --seed操作的随机种子。
* --swap-space每个 GPU 的 CPU 交换空间大小 (GiB)。

默认值：4

* --cpu-offload-gb每个 GPU 要卸载到 CPU 的空间，以 GiB 为单位。默认值为 0，表示不卸载。直观地看，此参数可以看作是增加 GPU 内存大小的虚拟方法。例如，如果您有一个 24 GB GPU 并将其设置为 10，则实际上可以将其视为 34 GB GPU。然后，您可以加载一个 13B 模型和 BF16 权重，这至少需要 26GB GPU 内存。请注意，这需要快速的 CPU-GPU 互连，因为模型的一部分在每次模型前向传递中都会从 CPU 内存动态加载到 GPU 内存。

默认值：0

* --gpu-memory-utilization用于模型执行器的 GPU 内存的比例，范围为 0 到 1。例如，值为 0.5 表示 50% 的 GPU 内存利用率。如果未指定，将使用默认值 0.9。这是一个按实例限制，仅适用于当前的 vLLM 实例。如果您在同一 GPU 上运行另一个 vLLM 实例，则无关紧要。例如，如果您在同一 GPU 上运行两个 vLLM 实例，则可以将每个实例的 GPU 内存利用率设置为 0.5。

默认值：0.9

* --num-gpu-blocks-override如果指定，则忽略 GPU 性能分析结果，并使用此 GPU 块数。用于测试抢占。
* --max-num-batched-tokens每次迭代的最大批处理 token 数。
* --max-num-partial-prefills对于分块预填充，最大并发部分预填充数。

默认值：1

* --max-long-partial-prefills对于分块预填充，将并发预填充的最大提示数（提示长度超过 –long-prefill-token-threshold）。将此值设置得小于 –max-num-partial-prefills 将允许较短的提示在某些情况下跳过较长提示的队列，从而提高延迟。

默认值：1

* --long-prefill-token-threshold对于分块预填充，如果提示长度超过此 token 数，则该请求被认为是长的。

默认值：0

* --max-num-seqs每次迭代的最大序列数。
* --max-logprobs要返回的最大对数概率数，logprobs 在 SamplingParams 中指定。

默认值：20

* --disable-log-stats禁用日志统计信息。

默认值：False

* --quantization, -q可选选项：aqlm, awq, deepspeedfp, tpu\_int8, fp8, ptpc\_fp8, fbgemm\_fp8, modelopt, nvfp4, marlin, gguf, gptq\_marlin\_24, gptq\_marlin, awq\_marlin, gptq, compressed-tensors, bitsandbytes, qqq, hqq, experts\_int8, neuron\_quant, ipex, quark, moe\_wna16, torchao, None

用于量化权重的技术。如果为 None，我们首先检查模型配置文件中的 quantization\_config 属性。如果该属性也为 None，我们则假定模型权重未被量化，并使用 dtype 来确定权重的data type（数据类型）。

* --rope-scalingRoPE scaling（RoPE 缩放）配置，JSON 格式。例如：<span class="pre">{"rope\_type":"dynamic","factor":2.0}</span>
* --rope-thetaRoPE theta。与 rope\_scaling 配合使用。在某些情况下，更改 RoPE theta 可以提高缩放模型的性能。
* --hf-token用作远程文件 HTTP bearer authorization（持有者授权）的token（令牌）。如果为 True，将使用运行 huggingface-cli login 时生成的token（存储在 ~/.huggingface 中）。
* --hf-overridesHuggingFace 配置的额外参数。这应该是一个 JSON 字符串，将被解析为一个字典。
* --enforce-eager始终使用 eager-mode（eager 模式） PyTorch。如果为 False，将结合使用 eager mode 和 CUDA graph（CUDA 图）以获得最大的性能和灵活性。

默认值：False

* --max-seq-len-to-captureCUDA graphs（CUDA 图）覆盖的最大序列长度。当序列的上下文长度大于此值时，我们将回退到 eager mode（eager 模式）。此外，对于 encoder-decoder（编码器-解码器）模型，如果 encoder input（编码器输入）的序列长度大于此值，我们也将回退到 eager mode（eager 模式）。

默认值：8192

* --disable-custom-all-reduce请参阅 ParallelConfig。

默认值：False

--tokenizer-pool-size用于异步 tokenization（分词）的 tokenizer pool（分词器池）的大小。如果为 0，将使用同步 tokenization（分词）。

默认值：0

--tokenizer-pool-type用于异步 tokenization（分词）的 tokenizer pool（分词器池）的类型。如果 tokenizer\_pool\_size 为 0，则忽略此项。

默认值：“ray”

--tokenizer-pool-extra-configtokenizer pool（分词器池）的额外配置。这应该是一个 JSON 字符串，将被解析为一个字典。如果 tokenizer\_pool\_size 为 0，则忽略此项。

--limit-mm-per-prompt对于每个 multimodal plugin（多模态插件），限制每个 prompt（提示）允许的输入实例数量。期望一个逗号分隔的项目列表，例如：image=16,video=2 允许每个 prompt（提示）最多 16 张图像和 2 个视频。默认为每种模态 1。

--mm-processor-kwargs多模态输入映射/处理的覆盖设置，例如，图像处理器。例如：。

--disable-mm-preprocessor-cache如果为 true，则禁用 multi-modal preprocessor/mapper（多模态预处理器/映射器）的缓存。（不推荐）

默认值：False

--enable-lora如果为 True，启用 LoRA adapters（LoRA 适配器）的处理。

默认值：False

--enable-lora-bias如果为 True，为 LoRA adapters（LoRA 适配器）启用 bias（偏置）。

默认值：False

--max-loras单个 batch（批次）中 LoRA 的最大数量。

默认值：1

--max-lora-rankLoRA rank（LoRA 秩）的最大值。

默认值：16

--lora-extra-vocab-sizeLoRA adapter（LoRA 适配器）中可能存在的额外 vocabulary（词汇表）的最大大小（添加到基础模型 vocabulary（词汇表））。

默认值：256

--lora-dtype可选选项：auto, float16, bfloat16

LoRA 的data type（数据类型）。如果为 auto，将默认为基础模型 dtype。

默认值：“auto”

--long-lora-scaling-factors指定多个 scaling factors（缩放因子）（可以与基础模型 scaling factor 不同 - 请参阅例如 Long LoRA），以允许同时使用使用这些 scaling factors 训练的多个 LoRA adapters（LoRA 适配器）。如果未指定，则仅允许使用使用基础模型 scaling factor 训练的 adapters（适配器）。

--max-cpu-loras存储在 CPU 内存中的 LoRA 的最大数量。必须 >= max\_loras。

--fully-sharded-loras默认情况下，只有一半的 LoRA 计算通过 tensor parallelism（张量并行）进行分片。启用此选项将使用 fully sharded layers（完全分片层）。在高序列长度、最大 rank（秩）或 tensor parallel size（张量并行大小）下，这可能更快。

默认值：False

--enable-prompt-adapter如果为 True，启用 PromptAdapters（Prompt 适配器）的处理。

默认值：False

--max-prompt-adapters一个 batch（批次）中 PromptAdapters（Prompt 适配器）的最大数量。

默认值：1

--max-prompt-adapter-tokenPromptAdapters tokens（Prompt 适配器令牌）的最大数量

默认值：0

--device可选选项：auto, cuda, neuron, cpu, tpu, xpu, hpu

vLLM 执行的 device type（设备类型）。

默认值：“auto”

--num-scheduler-steps每个 scheduler call（调度器调用）的最大 forward steps（前向步骤）数。

默认值：1

--use-tqdm-on-load, --no-use-tqdm-on-load加载模型权重时是否启用/禁用进度条。

默认值：True

--multi-step-stream-outputs如果为 False，则 multi-step（多步）将在所有步骤结束时 stream outputs（流式输出）

默认值：True

--scheduler-delay-factor在调度下一个 prompt（提示）之前，应用延迟（延迟因子乘以先前的 prompt latency（提示延迟））。

默认值：0.0

--enable-chunked-prefill如果设置，prefill requests（预填充请求）可以基于 max\_num\_batched\_tokens 进行分块。

--speculative-configspeculative decoding（推测解码）的配置。应为 JSON 字符串。

--model-loader-extra-configmodel loader（模型加载器）的额外配置。这将传递给与所选 load\_format（加载格式）对应的 model loader（模型加载器）。这应该是一个 JSON 字符串，将被解析为一个字典。

--ignore-patterns加载模型时要忽略的 pattern（模式）。默认值为 original/\*\*/\*，以避免重复加载 llama 的 checkpoints（检查点）。

默认值：[]

--preemption-mode如果为 ‘recompute’，引擎通过重新计算执行 preemption（抢占）；如果为 ‘swap’，引擎通过 block swapping（块交换）执行 preemption（抢占）。

--served-model-nameAPI 中使用的 model name(s)（模型名称）。如果提供了多个名称，服务器将响应任何提供的名称。响应的 model 字段中的模型名称将是列表中的第一个名称。如果未指定，模型名称将与 <span class="pre">--model</span> 参数相同。请注意，如果提供多个名称，此名称也将用于 prometheus metrics（普罗米修斯指标）的 model\_name tag content（标签内容）中，metrics tag（指标标签）将采用第一个名称。

--qlora-adapter-name-or-pathQLoRA adapter（QLoRA 适配器）的名称或路径。

--show-hidden-metrics-for-version启用自指定版本以来已隐藏的 deprecated（已弃用） Prometheus metrics（普罗米修斯指标）。例如，如果先前已弃用的 metric（指标）自 v0.7.0 版本以来已被隐藏，您可以使用 –show-hidden-metrics-for-version=0.7 作为临时应急方案，同时迁移到新 metrics（指标）。该 metric（指标）很可能在即将发布的版本中被完全删除。

--otlp-traces-endpointOpenTelemetry traces（OpenTelemetry 追踪）将被发送到的目标 URL。

--collect-detailed-traces有效选项为 model, worker, all。仅当设置了 <span class="pre">--otlp-traces-endpoint</span> 时，设置此项才有意义。如果设置，它将为指定的模块收集 detailed traces（详细追踪）。这涉及使用可能代价高昂和/或阻塞的操作，因此可能会对性能产生影响。

--disable-async-output-proc禁用 async output processing（异步输出处理）。这可能会导致性能降低。

默认值：False

--scheduling-policy可选选项：fcfs, priority

要使用的 scheduling policy（调度策略）。“fcfs”（先到先服务，即按照到达顺序处理请求；默认）或 “priority”（基于给定的 priority（优先级）处理请求（值越低表示越早处理），并以到达时间决定任何并列情况）。

默认值：“fcfs”

--scheduler-cls要使用的 scheduler class（调度器类）。“vllm.core.scheduler.Scheduler” 是默认的 scheduler（调度器）。可以是直接的类，也可以是 “mod.custom\_class” 形式的类路径。

默认值：“vllm.core.scheduler.Scheduler”

--override-neuron-config覆盖或设置 neuron device configuration（neuron 设备配置）。例如：。

--override-pooler-config覆盖或设置 pooling models（池化模型）的 pooling method（池化方法）。例如：。

--compilation-config, -O模型的 torch.compile 配置。当它是一个数字（0, 1, 2, 3）时，它将被解释为 optimization level（优化级别）。注意：级别 0 是默认级别，没有任何优化。级别 1 和 2 仅用于内部测试。级别 3 是推荐用于生产的级别。要指定完整的 compilation config（编译配置），请使用 JSON 字符串。按照传统编译器的惯例，也支持不带空格地使用 -O。-O3 等同于 -O 3。

--kv-transfer-configdistributed KV cache transfer（分布式 KV 缓存传输）的配置。应为 JSON 字符串。

--worker-cls用于 distributed execution（分布式执行）的 worker class（工作进程类）。

默认值：“auto”

--worker-extension-clsworker cls（工作进程类）之上的 worker extension class（工作进程扩展类），如果您只想向 worker class（工作进程类）添加新功能而不更改现有功能，这将非常有用。

默认值：“”

--generation-configgeneration config（生成配置）的文件夹路径。默认为 ‘auto’，generation config（生成配置）将从模型路径加载。如果设置为 ‘vllm’，则不加载 generation config（生成配置），将使用 vLLM 默认值。如果设置为文件夹路径，generation config（生成配置）将从指定的文件夹路径加载。如果在 generation config（生成配置）中指定了 max\_new\_tokens，则它将为所有请求设置服务器范围内的输出 tokens（令牌）数量限制。

默认值：auto

--override-generation-config以 JSON 格式覆盖或设置 generation config（生成配置）。例如：。如果与 –generation-config=auto 一起使用，override parameters（覆盖参数）将与模型的默认配置合并。如果 generation-config 为 None，则仅使用 override parameters（覆盖参数）。

--enable-sleep-mode为引擎启用 sleep mode（睡眠模式）。（仅支持 cuda 平台）

默认值：False

--calculate-kv-scales当 kv-cache-dtype 为 fp8 时，启用 k\_scale 和 v\_scale 的动态计算。如果 calculate-kv-scales 为 false，则 scales（缩放）将从模型 checkpoint（检查点）加载（如果可用）。否则，scales（缩放）将默认为 1.0。

默认值：False

--additional-configJSON 格式的指定平台的 additional config（附加配置）。不同的平台可能支持不同的配置。确保配置对于您正在使用的平台有效。输入格式类似于 ‘{“config\_key”:”config\_value”}’

--enable-reasoning是否为模型启用 reasoning\_content（推理内容）。如果启用，模型将能够生成 reasoning content（推理内容）。

默认值：False

--reasoning-parser可选选项：deepseek\_r1, granite

根据您正在使用的模型选择 reasoning parser（推理解析器）。这用于将 reasoning content（推理内容）解析为 OpenAI API 格式。 <span class="pre">--enable-reasoning</span> 需要此项。

--disable-cascade-attn为 V1 禁用 cascade attention（级联注意力）。虽然 cascade attention（级联注意力）不会改变数学上的正确性，但禁用它可以用于防止潜在的数值问题。请注意，即使将其设置为 False，cascade attention（级联注意力）也仅在 heuristic（启发式）表明它有利时才使用。

默认值：False

--disable-chunked-mm-input为 V1 禁用 multimodal input chunking attention（多模态输入分块注意力）。如果设置为 true 并且启用了 chunked prefill（分块预填充），我们不希望部分调度 multimodal item（多模态项目）。这确保了如果一个请求具有混合 prompt（提示）（例如文本 tokens TTTT 后跟图像 tokens IIIIIIIIII），其中只能调度一些图像 tokens（例如 TTTTIIIII，留下 IIIII），它将分步调度为 TTTT 和 IIIIIIIIII。

默认值：False

--disable-log-requests禁用 logging requests（请求日志记录）。

默认值：False

--max-log-len日志中打印的最大 prompt characters（提示字符）或 prompt ID numbers（提示 ID 号）数量。默认值 None 表示无限制。

--disable-fastapi-docs禁用 FastAPI 的 OpenAPI schema（OpenAPI 模式）、Swagger UI 和 ReDoc endpoint（ReDoc 端点）。

默认值：False

--enable-prompt-tokens-details如果设置为 True，则在 usage（用量）中启用 prompt\_tokens\_details。

默认值：False

--enable-server-load-tracking如果设置为 True，则在 app state（应用状态）中启用 tracking server\_load\_metrics（跟踪服务器负载指标）。

默认值：False

**Qwen3的部署命令和API调用**

|  |
| --- |
| Plain Text python -m vllm.entrypoints.openai.api\_server \  --model /root/autodl-tmp/models/Qwen/Qwen3-8B \  --served-model-name qwen3-8b \  --max-model-len 8k \  --host 0.0.0.0 \  --port 6006 \  --dtype bfloat16 \  --gpu-memory-utilization 0.8 \  --enable-auto-tool-choice \  --tool-call-parser hermes \     --enable-reasoning \  --reasoning-parser deepseek\_r1 \      python -m vllm.entrypoints.openai.api\_server \  --model /root/autodl-tmp/models/Qwen/Qwen2\_\_\_5-Omni-3B \  --served-model-name qwen-omni-3b \  --max-model-len 16k \  --host 0.0.0.0 \  --port 6006 \  --dtype float16 \  --gpu-memory-utilization 0.8      -- deepseek-r1-0528-qwen3-8B  python -m vllm.entrypoints.openai.api\_server \  --model /root/autodl-tmp/models/deepseek-ai/DeepSeek-R1-0528-Qwen3-8B \  --served-model-name ds-qwen3-8b \  --max-model-len 8k \  --host 0.0.0.0 \  --port 6006 \  --dtype bfloat16 \  --gpu-memory-utilization 0.8 \  --enable-auto-tool-choice \  --tool-call-parser hermes |

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/13.png)

**第二章、新版LangChain的应用开发**

开发环境： Pycharm-2025版+ Python-3.11 + JDK-17 + SpringAI(1.0.0-M7)- Spring-boot(3.44) langchain

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/14.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/15.png)

**架构**

LangChain作为一个框架由多个包组成。

* langchain-core

该包包含不同组件的基本抽象以及将它们组合在一起的方法。 核心组件的接口，如大型语言模型、向量存储、检索器等在此定义。 此处未定义任何第三方集成。 依赖项故意保持非常轻量级。

* langchain

主要的 langchain 包含链、代理和检索策略，这些构成了应用程序的认知架构。 这些不是第三方集成。 这里的所有链、代理和检索策略并不特定于任何一个集成，而是适用于所有集成的通用策略。

* langchain-community

此包包含由 LangChain 社区维护的第三方集成。 关键的合作伙伴包被单独列出（见下文）。 这包含了各种组件（大型语言模型、向量存储、检索器）的所有集成。 此包中的所有依赖项都是可选的，以保持包尽可能轻量。

**1、连接AI大模型 和 提示词模板**

小爱AI的注册地址：https://xiaoai.plus/register?aff=3TIp

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/16.png)

**一、连接AI大模型**

LangChain不托管任何聊天模型，而是依赖于第三方集成。官网如下：

https://www.langchain.com.cn/docs/integrations/chat/

在构建ChatModels时，我们有一些标准化参数：

* model: 模型名称
* temperature: 采样温度
* timeout: 请求超时
* max\_tokens: 生成的最大令牌数
* stop: 默认停止序列
* max\_retries: 请求重试的最大次数
* api\_key: 大模型供应商的API密钥
* base\_url: 发送请求的端点

一些重要事项需要注意：

* 标准参数仅适用于公开具有预期功能的参数的大模型供应商。例如，一些大模型供应商不公开最大输出令牌的配置，因此在这些大模型供应商上无法支持max\_tokens。
* 标准参数目前仅在具有自己集成包的集成上强制执行（例如 **langchain-openai**、langchain-anthropic 等），在 langchain-community 中的模型上不强制执行。

**二、提示词模板**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/17.png)

提示词模板有助于将用户输入和参数转换为语言模型的指令。 这可以用于指导模型的响应，帮助其理解上下文并生成相关且连贯的基于语言的输出。

提示词模板的输入是一个字典，其中每个键表示要填充的提示词模板中的变量。

有两种类型的提示词模板：

**字符串提示词模板**

这些提示词模板用于格式化单个字符串，通常用于更简单的输入。 例如，构造和使用 PromptTemplate 的一种常见方式如下：

|  |
| --- |
| Python prompt\_template = PromptTemplate.from\_template("帮我生成一个简短的，关于{topic}的报幕词。")  prompt\_template.invoke({"topic": "相声"}) |

[In-context Learning](https://zhida.zhihu.com/search?content_id=255359709&content_type=Article&match_order=1&q=In-context+Learning&zhida_source=entity)（ICL）作为一种新的自然语言处理范式逐渐崭露头角。ICL 的核心思想是：通过提供少量示例作为上下文，让大模型直接从中学习并做出预测。这一方法不仅省去了传统监督学习中繁琐的训练过程，还为大模型的应用开辟了新的可能性。

**聊天提示词模板**

这些提示词模板用于格式化消息列表。这些“模板”本身由一系列模板组成。 例如，构建和使用 ChatPromptTemplate 的一种常见方式如下：

|  |
| --- |
| Python from langchain\_core.prompts import ChatPromptTemplate  prompt\_template = ChatPromptTemplate.from\_messages([  ("system", "你是一个幽默的电视台主持人！"),  ("user", "帮我生成一个简短的，关于{topic}的报幕词。") ])  prompt\_template.invoke({"topic": "相声"}) |

在上述示例中，当调用此 ChatPromptTemplate 时，将构造两个消息。 第一个是系统消息，没有变量需要格式化。 第二个是 HumanMessage，将由用户传入的 topic 变量进行格式化。

**消息占位符**

此提示词模板负责在特定位置添加消息列表。 在上面的 ChatPromptTemplate 中，我们看到如何格式化两个消息，每个消息都是一个字符串。 但是如果我们希望用户传入一个消息列表（历史消息），并将其插入到特定位置呢？ 这就是需要使用 MessagesPlaceholder。

|  |
| --- |
| Python from langchain\_core.prompts import ChatPromptTemplate, MessagesPlaceholder from langchain\_core.messages import HumanMessage  prompt\_template = ChatPromptTemplate.from\_messages([  ("system", "你是一个电视台，高端访谈节目的主持人！"),  MessagesPlaceholder("msgs") ])  prompt\_template.invoke({"msgs": [HumanMessage(content="你好，主持人!")]}) |

这将生成一个包含两个消息的列表，第一个是系统消息，第二个是我们传入的 HumanMessage。后面的消息就是我 和AI 大模型对话过程中的历史消息。 这对于将消息列表插入到特定位置非常有用。

|  |
| --- |
| SQL prompt = ChatPromptTemplate.from\_messages([  ('system', '你是一个智能助手，尽可能的调用工具回答用户的问题'),  MessagesPlaceholder(variable\_name='chat\_history', optional=True),  ('human', '{input}'),  MessagesPlaceholder(variable\_name='agent\_scratchpad', optional=True), ]) |

**2、输出解析器和结构化输出**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/18.png)

输出解析器 ：负责获取模型的输出并将其转换为更适合下游任务的格式。 在使用大型语言模型生成结构化数据或规范化聊天模型和大型语言模型的输出时非常有用。

大型语言模型能够生成任意文本。这使得模型能够适当地响应广泛的 输入范围，但对于某些用例，限制大型语言模型的输出 为特定格式或结构是有用的。这被称为**结构化输出**。

例如，如果输出要存储在关系数据库中， 如果模型生成遵循定义的模式或格式的输出，将会容易得多。 最常见的输出格式将是JSON， 尽管其他格式如[YAML](https://www.langchain.com.cn/docs/how_to/output_parser_yaml/)也可能很有用。

**.with\_structured\_output()**

为了方便，一些LangChain聊天模型支持[.with\_structured\_output()](https://www.langchain.com.cn/docs/how_to/structured_output/#the-with_structured_output-method)方法。 该方法只需要一个模式作为输入，并返回一个字典或Pydantic对象。 通常，这个方法仅在支持下面描述的更高级方法的模型上存在， 并将在内部使用其中一种。它负责导入合适的输出解析器并 将模式格式化为模型所需的正确格式。

|  |
| --- |
| Python import json from typing import Optional  from langchain\_core.messages import HumanMessage from langchain\_core.output\_parsers import StrOutputParser from langchain\_core.prompts import MessagesPlaceholder, \  FewShotChatMessagePromptTemplate, PromptTemplate from pydantic import BaseModel, Field  from langchain\_demo.my\_llm import llm  from langchain\_core.prompts import ChatPromptTemplate   *# 生成一个笑话的段子： 三个属性，*  *# 使用pydantic定义一个类* class Joke(BaseModel):  *"""笑话（搞笑段子）的结构类(数据模型类 POVO)"""*  setup: str = Field(description="笑话的开头部分") *# 笑话的铺垫部分* punchline: str = Field(description="笑话的包袱/笑点") *# 笑话的爆笑部分* rating: Optional[int] = Field(description="笑话的有趣程度评分，范围1到10") *# 可选的笑话评分字段*   prompt\_template = PromptTemplate.from\_template("帮我生成一个关于{topic}的笑话。") runnable = llm.with\_structured\_output(Joke)  chain = prompt\_template | runnable resp = chain.invoke({"topic": "猫"}) print(resp)  print(resp.\_\_dict\_\_)  json\_str = json.dumps(resp.\_\_dict\_\_) print(json\_str) |

**SimpleJsonOutputParser**

一些模型，例如 [Mistral](https://www.langchain.com.cn/docs/integrations/chat/mistralai/)、[OpenAI](https://www.langchain.com.cn/docs/integrations/chat/openai/)， [Together AI](https://www.langchain.com.cn/docs/integrations/chat/together/) 和 [Ollama](https://www.langchain.com.cn/docs/integrations/chat/ollama/)， 支持一种称为 **JSON 模式** 的功能，通常通过配置启用。启用时，JSON 模式将限制模型的输出始终为某种有效的 JSON。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/19.png)

|  |
| --- |
| Python *# 创建聊天提示模板，要求模型以特定格式回答问题* prompt = ChatPromptTemplate.from\_template(  "尽你所能回答用户的问题。" *# 基本指令* '你必须始终输出一个包含"answer"和"followup\_question"键的JSON对象。其中"answer"代表：对用户问题的回答；"followup\_question"代表：用户可能提出的后续问题'  "{question}" *# 用户问题占位符* )   chain = prompt | llm | SimpleJsonOutputParser() resp = chain.invoke({"question": "细胞的动力源是什么？"}) print(resp) |

**工具调用**

对于支持此功能的模型，工具调用 可以非常方便地生成结构化输出。它消除了 关于如何最好地提示模式的猜测，而是采用内置模型功能。

它的工作原理是首先将所需的模式直接或通过 LangChain 工具 绑定到 聊天模型，使用 .bind\_tools() 方法。然后模型将生成一个包含 与所需形状匹配的 args 的 tool\_calls 字段的 AIMessage。工具调用是一种通常一致的方法，可以让模型生成结构化输出，并且是默认技术 用于 [.with\_structured\_output()](https://www.langchain.com.cn/docs/concepts/#with_structured_output) 方法，当模型支持时。

|  |
| --- |
| Python class ResponseFormatter(BaseModel):  """始终使用此工具来结构化你的用户响应""" *# 文档字符串说明这个类用于格式化响应*  answer: str = Field(description="对用户问题的回答") *# 回答内容字段* followup\_question: str = Field(description="用户可能提出的后续问题") *# 后续问题字段*    runnable = llm.bind\_tools([ResponseFormatter])  resp = runnable.invoke("细胞的动力源是什么？") print(resp.tool\_calls[-1]['args']) resp.pretty\_print() |

**3、大模型应用开发案例**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/20.png)

**一、多模态的聊天机器人**

1. 多模型调用和多模态模型的调用
2. 保存历史聊天记录
3. 修剪聊天上下文
4. 形成摘要记忆
5. 拥有web界面，方便用户使用
6. 可以在线录制语音
7. 可以处理语音、图片和视频

**注意： 目前所有的多模态大模型，如果需要传入多媒体内容。只有两种方式：**

1、传入多媒体文件的网络访问路径，比如: http://www..baidu.com/log.png

2、把多媒体文件转换为base64格式的字符串，再传入大模型。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/21.png)

**4、RAG和Embeddings模型**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/22.png)

**一、什么是Embedding模型**

Embedding模型是指将高维度的数据（例如文字、图片、视频）映射到低维度空间的过程。简单来说，embedding向量就是一个N维的实值向量，它将输入的数据表示成一个连续的数值空间中的点。

Embeddings的学习通常基于无监督或弱监督的方法。对于自然语言处理任务，常用的Embeddings方法包括Word2Vec、GloVe和FastText。这些方法可以从大规模的文本语料库中学习单词的分布式表示。对于计算机视觉任务，常用的Embeddings方法包括卷积神经网络（CNN）和循环神经网络（RNN）等。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/23.png)

通俗易懂的描述：嵌入就相当于给文本穿上了“数字化”的外衣，目的是让机器更好的理解和处理。

**发展：**由静态的Word Embedding（如Word2Vec、GloVe和FastText） -> 动态预训练模型（如ELMo、BERT、GPT、GPT-2、GPT-3、ALBERT、XLNet等）。大型语言模型可以生成上下文相关的 embedding 表示，可以更好地捕捉单词的语义和上下文信息。

**向量空间（Vector Space）**

所有的数据都变成向量，这些向量组成一个庞大的矩阵。在这个世界里，每个词、句子、图片、用户…都被表示成一个“点”（即向量），大家都有自己的“坐标”。

我们可以通过“距离”和“方向”来理解它们的关系。

Embedding 向量放在向量空间里，有啥用？

距离表示相似度

向量之间越近：意义越相似

向量之间越远：意义越不同

比如：

“苹果 🍎” 和 “香蕉 🍌” 的向量夹角小（近） → 都是水果

“苹果 🍎” 和 “MacBook 💻” 的向量略远 → 一个是水果，一个是电子产品

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/24.png)

**使用场景**

1. Embeddings可以在各种机器学习任务中使用，包括分类、聚类、检索和推荐等。
2. 在自然语言处理任务中，可以使用静态预训练的Embeddings模型，如Word2Vec、GloVe和FastText，来生成单词的向量表示。这些预训练的Embeddings模型通常在大规模的文本数据上进行训练，可以用于处理不同的自然语言处理任务，如情感分析、命名实体识别和机器翻译等。
3. 在计算机视觉任务中，可以使用卷积神经网络（CNN）提取图像的特征向量，然后使用这些特征向量进行分类、检索和生成等任务。另外，通过将图像与文本进行联合训练，可以学习到图像和文本之间的语义关系，从而实现图像与文本的检索和生成等任务。
4. 我们在做RGA开发时又会涉及到向量数据库，在创建向量数据库时又需要使用Embedding模型对文本进行向量化处理。在检索的时候，需要对用户输入进行向量化处理也需要用到。

**解决问题**

* 降维：在高维度空间中，数据点之间可能存在很大的距离，使得样本稀疏，嵌入模型可以减少数据稀疏性。
* 捕捉语义信息：Embedding不仅仅是降维，更重要的是，它能够捕捉到数据的语义信息。语义相近的词在向量上也是相近的
* 特征表示：原始数据的特征往往难以直接使用，通过嵌入模型可以将特征转换成更有意义的表示。
* 计算效率：在低维度空间中对数据进行处理和分析往往更加高效。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/25.png)

**独热编码（One-Hot Encoding）**

是一种将数据转换为二进制向量的技术。它的主要目的是将分类变量转换为机器学习算法能够处理的格式，从而避免数值关系的误判。

举例：词表中有 10,000 个词，每个词都用一个只有一个 1，其它全是 0 的向量来表示。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/26.png)

这样的向量是：

📏 高维（非常长，比如 10k、100k…）

⚪ 稀疏（只有一个 1，其他都是 0）

🧱 没有语义信息（“猫”和“狗”之间毫无关系）

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/27.png)

**二、langchain的文本嵌入模型(Embeddings)**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/28.png)

嵌入模型创建文本片段的向量表示。您可以将向量视为一个数字数组，它捕捉了文本的语义含义。 通过这种方式表示文本，您可以执行数学运算，从而进行诸如搜索其他在意义上最相似的文本等操作。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/29.png)

Embeddings 类是一个用于与文本嵌入模型接口的类。有很多嵌入大模型供应商（OpenAI、Hugging Face，BGE 等） - 这个类旨在为它们提供一个标准接口。

LangChain 中的基础 Embeddings 类提供了两个方法：一个用于嵌入文档，一个用于嵌入查询。前者，.embed\_documents，接受多个文本作为输入，而后者，.embed\_query，接受单个文本。将这两个方法分开是因为某些嵌入大模型供应商对文档（待搜索的内容）和查询（搜索查询本身）有不同的嵌入方法。 .embed\_query 将返回一个浮点数列表，而 .embed\_documents 返回一个浮点数列表的列表。

**三、私有化部署Qwen3-Embedding**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/30.png)

**生产环境中的Embeddings模型对比**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/31.png)

**点击图片可查看完整电子表格**

为了部署Embedding模型，我们需要引入对应的工具库，目前主要有几类：

1. **Sentence-Transformers**: Sentence-Transformers库是基于HuggingFace的Transformers库构建的，它专门设计用于生成句子级别的嵌入。它引入了一些特定的模型和池化技术，使得生成的嵌入能够更好地捕捉句子的语义信息。Sentence-Transformers库特别适合于需要计算句子相似度、进行语义搜索和挖掘同义词等任务。
2. **HuggingFace Transformers**: HuggingFace的Transformers库是一个广泛使用的NLP库，它提供了多种预训练模型，如BERT、GPT-2、RoBERTa等。这些模型可以应用于各种NLP任务，如文本分类、命名实体识别、问答系统等。Transformers库支持多种编程语言，并且支持模型的微调和自定义模型的创建。虽然Transformers库的功能强大，但它主要关注于模型的使用，而不是直接提供句子级别的嵌入。
3. **Langchain集成**的HuggingFaceBgeEmbeddings。与3一样。
4. **FlagEmbedding:** 这是一个相对较新的库，其核心在于能够将任意文本映射到低维稠密向量空间，以便于后续的检索、分类、聚类或语义匹配等任务。FlagEmbedding的一大特色是它可以支持为大模型调用外部知识，这意味着它不仅可以处理纯文本数据，还能整合其他类型的信息源，如知识图谱等，以提供更丰富的语义表示。

总的来说，FlagEmbedding强调的是稠密向量的生成和外部知识的融合；HuggingFace Transformers提供了一个广泛的预训练模型集合，适用于多种NLP任务；而Sentence-Transformers则专注于生成高质量的句子嵌入，适合那些需要深入理解句子语义的应用场景。

**四、BGE-Large的Embadding+Huggingface私有化**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/32.png)

HuggingFace 上的 BGE 模型是最好的开源嵌入模型之一。 BGE 模型由北京人工智能研究院 （BAAI） 创建。 是一家从事 AI 研发的私营非营利组织。

**BGE-Large**（智源研究院）和GTE-Large（阿里巴巴）（**6月之前**）是当前中文RAG领域主流的开源Embedding模型，两者的核心区别与优势如下：

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/33.png)

**配置HuggingFace镜像站：**[**https://hf-mirror.com/**](https://hf-mirror.com/)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/34.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/35.png)

|  |
| --- |
| Plain Text pip install --upgrade --quiet sentence\_transformers |

|  |
| --- |
| Plain Text from langchain.embeddings import HuggingFaceBgeEmbeddings model\_name = "BAAI/bge-large-zh-v1.5" model\_kwargs = {'device': 'cuda'} encode\_kwargs = {'normalize\_embeddings': True} # set True to compute cosine similarity model = HuggingFaceBgeEmbeddings(  model\_name=model\_name,  model\_kwargs=model\_kwargs,  encode\_kwargs=encode\_kwargs,  query\_instruction="为这个句子生成表示以用于检索相关文章：" ) model.query\_instruction = "为这个句子生成表示以用于检索相关文章：" |

在 HuggingFaceBgeEmbeddings 中，normalize\_embeddings 参数通常只接受布尔值（True 或 False），用于决定是否对生成的嵌入向量进行归一化处理。具体来说：

* **True**：生成的嵌入向量会被归一化为单位向量。这意味着每个嵌入向量的 L2 范数（欧几里得长度）将被缩放到 1。
* **False**：生成的嵌入向量将保持原始的数值，不进行归一化处理。
* **优点** ：
* **提高相似度计算的稳定性** ：在许多应用场景中，如余弦相似度计算，归一化后的向量可以避免因向量长度不同而导致的相似度偏差，使相似度计算更加专注于向量的方向而非长度。
* **一致性** ：在某些情况下，归一化可以确保不同批次或不同模型生成的嵌入向量在同一尺度上，便于比较和整合。

**案例：根据语义搜索美食评论数据**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/36.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/37.png)

**余弦距离（Cosine Distance）​**的计算，用于衡量两个向量在方向上的相似性。代表 文本语义的相似性

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/38.png)

||a|| ： 计算向量 a 的欧几里得范数（L2范数），即向量的长度。公式为 sqrt(a₁² + a₂² + ... + aₙ²)

ab： 计算向量 a 和 b 的点积（内积），即对应元素相乘后求和。

**五、向量数据库**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/39.png)

存储和搜索非结构化数据的最常见方法之一是将其嵌入并存储生成的嵌入向量， 然后在查询时嵌入非结构化查询并检索与嵌入查询“最相似”的嵌入向量。 向量存储负责存储嵌入数据并执行向量搜索， 为您处理这些。

**Chroma**

**Chroma** 是一个开源的向量数据库，专注于简化文本嵌入的存储和检索过程。Chroma 采用 Apache 2.0 许可证。它的主要特点包括：

1. **支持多种存储后端**：Chroma支持多种底层存储选项，如DuckDB（适用于独立应用）和ClickHouse（适用于大规模扩展）。
2. **多语言支持**：Chroma提供了Python和JavaScript/TypeScript的SDK，方便开发者快速集成。
3. **简单易用**：Chroma的设计理念是“简单至上”，旨在提升开发者的效率。
4. **高性能**：Chroma不仅支持快速的相似度搜索，还提供了对搜索结果的分析功能。

**FAISS**

Faiss是由Facebook AI Research团队开发的一个库，旨在高效地进行大规模向量相似性搜索。它不仅支持CPU，还能利用GPU进行加速，非常适合处理大量高维数据。Faiss提供了多种索引类型，以适应不同的需求，从简单的平面索引（Flat Index）到更复杂的倒排文件索引（IVF）和乘积量化索引（PQ）。

**Milvus**

**Milvus 基本介绍**

* Milvus 由 Zilliz 开发，并捐赠给了 Linux 基金会下的 LF AI & Data 基金会，现已成为世界领先的开源向量数据库项目之一
* 什么是向量数据库：传统的数据库主要处理结构化数据，而向量数据库则专注于处理非结构化数据经过嵌入模型（embedding model）转换而来的向量数据。这些向量是高维空间中的点，它们捕获了原始数据的语义信息。向量数据库的核心能力是进行相似性搜索，即根据查询向量找到最相似的向量，从而实现语义级别的搜索和匹配。
* Milvus 采用 Apache 2.0 许可发布，大多数贡献者都是高性能计算（HPC）领域的专家，擅长构建大规模系统和优化硬件感知代码

**六、RAG案例**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/40.png)

RAG：Retrieval-Augmented Generation 检索增强生成。RAG通过结合LLMs的内在知识和外部数据库的非参数化数据，提高了模型在知识密集型任务中的准确性和可信度。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/41.jpeg)

**上下文感知的检索器**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/42.png)

**第三章、基于LangGraph的Agent**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/43.png)

LangGraph 专为希望构建强大、适应性强的 AI 智能体的开发者而设计。开发者选择 LangGraph 的原因是：

* **可靠性和可控性。**通过审核检查和人工干预审批来指导智能体行为。LangGraph 可为长时间运行的工作流持久化上下文，使您的智能体保持正常运行。
* **低层级和可扩展性。**使用完全描述性的低层级原语构建自定义智能体，不受限制自定义的僵化抽象约束。设计可扩展的多智能体系统，其中每个智能体都为您的用例量身定制特定角色。
* **一流的流式传输支持。**通过逐令牌流式传输和中间步骤流式传输，LangGraph 让用户实时清晰地了解智能体的推理和行动过程。

LangGraph支持两种对于构建对话代理至关重要的内存类型：

* **短期内存**：通过在会话中维护消息历史来跟踪正在进行的对话。
* **长期内存**：在不同会话之间存储用户特定或应用程序级别的数据。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/44.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/45.png)

**什么是Agent？**

人类在复杂的模式识别任务中表现卓越，但通常需要借助工具（如书籍、搜索引擎或计算器）来补充先验知识以得出结论。同理，生成式AI模型可通过训练使用工具获取实时信息或建议的实际动作。例如：

* 模型可利用数据库检索工具获取客户购买历史以生成个性化购物推荐
* 基于用户查询，模型可通过API调用发送邮件或完成金融交易

为实现此能力，模型需具备：

1. 外部工具集访问权限
2. 自主规划与执行任务的推理能力

这种结合推理逻辑与外部信息访问的系统，即构成智能体（Agent）。

智能体的认知架构中有三个基本组件：模型（Model）、工具（Tools）和以及一个提供指令的**提示**。

LLM 在一个循环中运行。在每次迭代中，它会选择一个要调用的工具，提供输入，接收结果（一个观察），并利用该观察来指导下一个动作。循环会一直持续，**直到满足停止条件——通常是Agent已经收集到足够的信息来响应用户时。**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/46.png)

**Agent vs. Workflow（图）**

Anthropic将Agent系统划分为两类：

1. 第一类是workflow。遵循预定义的工作流，编排LLM和工具，固定代码路径。
2. Agent：此类Agent被定义为完全自主的系统，这些系统在较长时间内独立运行，可以动态地指导自身流程和工具使用的系统。通过自身的推理、规划能力，自主控制，完成任务。

**注意：在LangGraph中，一切都是图，Agent是图中的一个节点。使用** [**create\_react\_agent**](https://github.langchain.ac.cn/langgraph/reference/agents/#langgraph.prebuilt.chat_agent_executor.create_react_agent) **创建一个Agent，并且也得到了一个图。**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/47.png)

以下代码片段展示了如何使用 [create\_react\_agent](https://github.langchain.ac.cn/langgraph/reference/agents/#langgraph.prebuilt.chat_agent_executor.create_react_agent) 创建上述Agent

|  |
| --- |
| Python from langgraph.prebuilt import create\_react\_agent from langchain\_openai import ChatOpenAI  model = ChatOpenAI("xxx")  def tool() -> None:  """Testing tool."""  ...  agent = create\_react\_agent(  model,  tools=[tool], )  agent.get\_graph().draw\_mermaid\_png() |

**1、LangGraph本地服务**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/48.png)

**LangGraph CLI** 是一个多平台命令行工具，用于在本地构建和运行 LangGraph API 服务器。生成的服务器包含您的图的所有运行、线程、助手等的 API 端点，以及运行您的代理所需的其他服务，包括用于检查点和存储的托管数据库。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/49.png)

**安装和使用步骤**

**一、创建Python虚拟环境**

虚拟环境的安装步骤

1. 安装好python解释器： Python >= 3.11 is required.
2. 安装虚拟环境库，在cmd中输入：

|  |
| --- |
| Plain Text pip install virtualenv |

1. 创建虚拟环境，在cmd中切换到需要创建虚拟环境的目录下，执行：

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/50.png)

|  |
| --- |
| Plain Text virtualenv env\_name |

1. 激活虚拟环境，在cmd中进入到 第三步创建的 env\_name/Scripts 目录下，执行：

|  |
| --- |
| Plain Text activate |

执行成功后，在cmd中，当前输入行前面会有 (env\_name) 的前缀

在当前状态下，使用 pip 就是在虚拟环境中安装第三方库了

5. 退出虚拟环境，cmd中输入：

|  |
| --- |
| Plain Text deactivate |

**二、安装LangGraph CLI**

|  |
| --- |
| Python # Python >= 3.11 is required.  pip install --upgrade "langgraph-cli[inmem]" |

**三、创建 LangGraph 应用**

从 [new-langgraph-project-python 模板](https://github.com/langchain-ai/new-langgraph-project)或 [new-langgraph-project-js 模板](https://github.com/langchain-ai/new-langgraphjs-project)创建一个新应用。此模板演示了一个单节点应用，您可以根据自己的逻辑进行扩展。

**注意:如果您使用 langgraph new 命令时未指定模板，将显示一个交互式菜单，允许您从可用模板列表中进行选择。**

|  |
| --- |
| Python langgraph new path/to/your/app --template new-langgraph-project-python |

**四、安装项目依赖**

在您的新 LangGraph 应用的根目录下，以编辑模式安装依赖项，以便服务器使用您的本地更改。

在 **LangGraph** 中，pyproject.toml 代传统的 setup.py 和 requirements.txt.可能包含以下扩展配置：

* **依赖分组**​：如 [project.optional-dependencies] 定义 dev（开发工具）和 test（测试框架）依赖。
* **动态版本控制**​：通过 requires-python = ">=3.9" 指定 Python 版本兼容性。
* **CI/CD 集成**​：通过 [tool.\*] 配置与 GitHub Actions 或 GitLab CI 的交互

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/51.png)

|  |
| --- |
| Python cd path/to/your/app pip install -e . |

**五、修改graph.py的代码**

|  |
| --- |
| Python *# 本地私有化部署的大模型* llm = ChatOpenAI(  model='qwen3-8b',  temperature=0.8,  api\_key='xx',  base\_url="http://localhost:6006/v1",  extra\_body={'chat\_template\_kwargs': {'enable\_thinking': False}}, )  def get\_weather(city: str) -> str:  *"""Get weather for a given city."""* return f"It's always sunny in {city}!"  graph = create\_react\_agent(  llm,  tools=[get\_weather],  prompt="You are a helpful assistant" ) |

**六、启动LangGraph服务器**

langgraph dev 命令以内存模式启动 LangGraph 服务器。此模式适用于开发和测试目的。对于生产用途，请部署 LangGraph 服务器并使其能够访问持久存储后端。

|  |
| --- |
| Python 命令： langgraph dev |

**选项**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/52.png)

**点击图片可查看完整电子表格**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/53.png)

**七、测试和访问agent的API**

1. LangGraph Studio 是一个专用 UI，您可以将其连接到 LangGraph API 服务器，以在本地可视化、交互和调试您的应用。通过访问 langgraph dev 命令输出中提供的 URL，在 LangGraph Studio 中测试您的Agent和图。
2. PythonSDK测试

|  |
| --- |
| Shell pip install langgraph-sdk |

异步测试：

|  |
| --- |
| Python from langgraph\_sdk import get\_client import asyncio  client = get\_client(url="http://localhost:2024")  async def main():  async for chunk in client.runs.stream(  None, # Threadless run  "agent", # Name of assistant. Defined in langgraph.json.  input={  "messages": [{  "role": "human",  "content": "What is LangGraph?",  }],  },  ):  print(f"Receiving new event of type: {chunk.event}...")  print(chunk.data)  print("\n\n")  asyncio.run(main()) |

同步测试：

|  |
| --- |
| Python from langgraph\_sdk import get\_sync\_client  client = get\_sync\_client(url="http://localhost:2024")  for chunk in client.runs.stream(  None, # Threadless run  "agent", # Name of assistant. Defined in langgraph.json.  input={  "messages": [{  "role": "human",  "content": "What is LangGraph?",  }],  },  stream\_mode="messages-tuple", ):  print(f"Receiving new event of type: {chunk.event}...")  print(chunk.data)  print("\n\n") |

1. JavaScript SDK测试

安装 LangGraph JS SDK

|  |
| --- |
| Shell npm install @langchain/langgraph-sdk |

向LangGraph服务区发送消息：

|  |
| --- |
| JavaScript const { Client } = await import("@langchain/langgraph-sdk");  // only set the apiUrl if you changed the default port when calling langgraph dev const client = new Client({ apiUrl: "http://localhost:2024"});  const streamResponse = client.runs.stream(  null, // Threadless run  "agent", // Assistant ID  {  input: {  "messages": [  { "role": "user", "content": "What is LangGraph?"}  ]  },  streamMode: "messages-tuple",  } );  for await (const chunk of streamResponse) {  console.log(`Receiving new event of type: ${chunk.event}...`);  console.log(JSON.stringify(chunk.data));  console.log("\n\n"); } |

1. REST API测试：

|  |
| --- |
| Shell curl -s --request POST \  --url "http://localhost:2024/runs/stream" \  --header 'Content-Type: application/json' \  --data "{  \"assistant\_id\": \"agent\",  \"input\": {  \"messages\": [  {  \"role\": \"human\",  \"content\": \"What is LangGraph?\"  }  ]  },  \"stream\_mode\": \"messages-tuple\"  }" |

**2、Tool工具的定义**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/54.png)

在构建Agent 时，您需要为其提供一个它可以使用的 工具 列表。除了实际调用的函数之外，工具还包括几个组件：

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/55.png)

**注意： 如果工具具有精心选择的名称、描述和args\_schema，模型将表现得更好。**

LangChain 支持从以下对象创建工具

1. 函数；
2. LangChain Runnables；
3. 通过从 BaseTool 子类化 -- 这是最灵活的方法，它提供了最大的控制程度，但代价是需要付出更多的努力和编写更多的代码。

**一、从函数创建工具**

这个 @tool 装饰器是定义自定义工具的最简单方法。默认情况下，装饰器使用函数名称作为工具名称，但可以通过将字符串作为第一个参数传递来覆盖。此外，装饰器将使用函数的文档字符串作为工具的描述 - 因此必须提供文档字符串。请注意，@tool 支持解析注释、嵌套模式和其他特性：

**sglang部署Qwen-3-8B模型**

|  |
| --- |
| SQL # 安装sglang pip install "sglang[all]>=0.4.6.post1"  # 启动大模型服务 python -m sglang.launch\_server \  --model-path /hy-tmp/models/Qwen/Qwen3-8B \  --served-model-name qwen3-8b \  --context-length 8192 \  --trust-remote-code \  --host 0.0.0.0 \  --port 8080 \  --reasoning-parser qwen3 \  --tool-call-parser qwen25 |

**二、从可运行对象（Runnable）创建工具**

接受字符串或 dict 输入的 LangChain [Runnables](https://www.langchain.com.cn/docs/concepts/#runnable-interface) 可以使用 [as\_tool](https://python.langchain.com/api_reference/core/runnables/langchain_core.runnables.base.Runnable.html#langchain_core.runnables.base.Runnable.as_tool) 方法转换为工具，该方法允许为参数指定名称、描述和其他模式信息。

**三、子类化 BaseTool**

可以通过从 BaseTool 子类化来定义自定义工具。这提供了对工具定义的最大控制，但需要编写更多代码。

|  |
| --- |
| Python class SearchArgs(BaseModel):  query: str = Field(description="需要进行网络搜索的信息。")   *# 网络搜索的工具* class MySearchTool(BaseTool):  *# 工具名字* name: str = "search\_tool"   description: str = '搜索互联网上公开内容的工具'   return\_direct: bool = False   args\_schema: Type[BaseModel] = SearchArgs   def \_run(self, query) -> str:  try:  print("执行我的Python中的工具，输入的参数为:", query)  response = zhipuai\_client.web\_search.web\_search(  search\_engine="search\_pro",  search\_query=query  )  *# print(response)* if response.search\_result:  return "\n\n".join([d.content for d in response.search\_result])  return '没有搜索到任何内容！'  except Exception as e:  print(e)  return '没有搜索到任何内容！' |

**3、Agent的上下文和记忆**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/56.png)

**一、上下文**

小爱AI的注册地址：https://xiaoai.plus/register?aff=3TIp

上下文包括消息列表之外的*任何*数据，这些数据可以影响代理行为或工具执行。这可以是：

* 运行时传入的信息，如 `user\_id` 或 API 凭据。
* 多步推理过程中更新的内部状态。
* 来自先前交互的持久记忆或事实。

LangGraph 提供了**三种**提供上下文的主要方式：

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/57.png)

**点击图片可查看完整电子表格**

您可以使用上下文来：

* 调整模型看到的系统提示
* 为工具提供必要的输入
* 在正在进行的对话中跟踪事实

**Configurable**

配置适用于不可变数据，如用户元数据或 API 密钥。当您有在运行期间不会更改的值时使用。

使用保留用于此目的的键 **"configurable"** 指定配置。

|  |
| --- |
| C++ agent.invoke(  {"messages": [{"role": "user", "content": "hi!"}]},  config={"configurable": {"user\_id": "user\_123"}} ) |

**案例：通过Configurable中掺入参数，来动态设置系统提示词**

**状态 AgentState(可变上下文)**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/58.png)

状态在运行期间充当Agent的记忆，可以短期存储也可以长期存储。它保存可在执行期间演变的动态数据，例如从工具或 LLM 输出派生的值。

1、搞清楚AgentState的作用。

2、案例：（给用户发出一个祝福语句）输入username ---> config----> 工具1---> 把username修改到State中------> 工具2----->获取State的username得到最终答案。

**二、记忆存储**

**这是一个强大的功能，允许您在多次调用中持久化代理的状态。否则，状态仅限于单次运行。**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/59.png)

**短期存储：线程级存储（会话级）**

短期存储使Agent能够跟踪多轮对话。要使用它，您必须：

1. 在创建代理时提供checkpointer。checkpointer可以实现代理状态的持久性。
2. 在运行代理时在配置中提供thread\_id。thread\_id是对话会话的唯一标识符。

|  |
| --- |
| Python # 内存： 开发环境 checkpointer = InMemorySaver()  # 在生产环境中，使用由数据库支持的检查点 DB\_URI = "postgresql://postgres:postgres@localhost:5432/postgres?sslmode=disable" with PostgresSaver.from\_conn\_string(DB\_URI) as checkpointer: # 必须安装：pip install -U "psycopg[binary,pool]" langgraph-checkpoint-postgres  # 生产环境：Redis # pip install -U langgraph-checkpoint-redis DB\_URI = "redis://:6379" with RedisSaver.from\_conn\_string(DB\_URI) as checkpointer: |

**长期存储：跨线程存储**

使用长期内存来跨会话存储用户特定或应用程序特定的数据。这对于聊天机器人等应用程序非常有用，您可能希望记住用户偏好或其他信息。

要使用长期内存，您需要：

1. 配置一个存储以在调用之间持久化数据。
2. 使用get\_store函数从工具或提示中访问存储。

|  |
| --- |
| Python # 开发环境中： 内存 store = InMemoryStore()  # 在生产环境中，使用由数据库支持的存储 DB\_URI = "postgresql://postgres:postgres@localhost:5442/postgres?sslmode=disable" with (  PostgresStore.from\_conn\_string(DB\_URI) as store,   PostgresSaver.from\_conn\_string(DB\_URI) as checkpointer, ):   # store.setup()   # checkpointer.setup()    DB\_URI = "redis://:6379" with (   RedisStore.from\_conn\_string(DB\_URI) as store,   RedisSaver.from\_conn\_string(DB\_URI) as checkpointer, ):   store.setup()   checkpointer.setup() |

**注意：**

* 首次使用 Postgres 存储时，您需要调用 store.setup()，checkpointer.setup()
* 首次使用 Redis 存储时，您需要调用 store.setup()，checkpointer.setup()

**第四章、Agent和 MCP开发**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/60.png)

**1、Function Calling（函数调用）**

**Function Calling**由OpenAI等公司推动，允许大语言模型与外部工具连接，将自然语言转换为API调用。这解决了大模型在训练结束，就知识更新停滞的问题。

通过调用外部工具和服务，Function Calling帮助大模型解决了比如今天温度多少度，今天的大盘收盘点数是多少之类的实时性问题。

Function Calling的工作原理可以通过以下几个简单步骤来理解：

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/61.png)

以天气查询为例，Function Calling的工作流程大致如下：

**第一步，识别需求** ：这是一个关于实时天气的问题，需要调用外部天气API。

**第二步，选择函数** ：从可用函数中选择get\_current\_weather函数。

**第三步，准备参数** ：

|  |
| --- |
| Plain Text {  "location": "北京",  "unit": "celsius" } |

**第四步，调用函数** ：系统使用这些参数调用实际的天气API，获取北京的实时天气数据。

**第五步，整合回答** ： "根据最新数据，北京今天的天气晴朗，当前温度23°C，湿度45%，微风。今天的最高温度预计为26°C，最低温度为18°C。"

对开发者来说，使用LLM的Function Calling起步相对容易，只需按照API要求定义函数规格（通常JSON模式）并将其随请求发送，模型就可能按照需要调用这些函数，逻辑较直观。

因此，对于\*\*单一模型、少量功能\*\*的简单应用，Function Calling 实现起来非常直接，几乎“一键”将模型输出对接到代码逻辑中。

然而，它的\*\*局限\*\*在于缺乏跨模型的一致性：每个LLM供应商的接口格式略有差异，开发者若想支持多个模型，需要为不同API做适配或使用额外框架处理。而且，\*\*无状态性\*\* ：模型仅生成调用规范，实际执行由外部系统完成。

**2、MCP协议和开发**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/62.png)

**1、定义**

MCP（Model Context Protocol，模型上下文协议） ，2024年11月底，由 Anthropic 推出的一种开放标准，旨在统一大型语言模型（LLM）与外部数据源和工具之间的通信协议。MCP 的主要目的在于解决当前 AI 模型因数据孤岛限制而无法充分发挥潜力的难题，MCP 使得 AI 应用能够安全地访问和操作本地及远程数据，为 AI 应用提供了连接万物的接口。

|  |
| --- |
| Function Calling是AI模型调用函数的机制，MCP是一个标准协议，使AI模型与API无缝交互，而AI Agent是一个自主运行的智能系统，利用Function Calling和MCP来分析和执行任务，实现特定目标。 |

即使是最强大模型也会受到数据隔离的限制，形成信息孤岛，要做出更强大的模型，每个新数据源都需要自己重新定制实现，使真正互联的系统难以扩展，存在很多的局限性。

现在，MCP 可以直接在 AI 与数据（包括本地数据和互联网数据）之间架起一座桥梁，通过 MCP 服务器和 MCP 客户端，大家只要都遵循这套协议，就能实现“万物互联”。

有了MCP，可以和数据和文件系统、开发工具、Web 和浏览器自动化、生产力和通信、各种社区生态能力全部集成，实现强大的协作工作能力，它的价值远不可估量。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/63.png)

**2、MCP 与 Function Calling 的区别**

* MCP（Model Context Protocol），模型上下文协议
* Function Calling，函数调用

这两种技术都旨在增强 AI 模型与外部数据的交互能力，但 MCP 不止可以增强 AI 模型，还可以是其他的应用系统。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/64.png)

**3、工作原理**

MCP 协议采用了一种独特的架构设计，它将 LLM 与资源之间的通信划分为三个主要部分：客户端、服务器和资源。

客户端负责发送请求给 MCP 服务器，服务器则将这些请求转发给相应的资源。这种分层的设计使得 MCP 协议能够更好地控制访问权限，确保只有经过授权的用户才能访问特定的资源。

以下是 MCP 的基本工作流程：

* 初始化连接：客户端向服务器发送连接请求，建立通信通道。
* 发送请求：客户端根据需求构建请求消息，并发送给服务器。
* 处理请求：服务器接收到请求后，解析请求内容，执行相应的操作（如查询数据库、读取文件等）。
* 返回结果：服务器将处理结果封装成响应消息，发送回客户端。
* 断开连接：任务完成后，客户端可以主动关闭连接或等待服务器超时关闭。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/65.png)

**4、通信机制**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/66.png)

MCP 协议支持三种主要的通信机制：

* 1、基于标准输入输出的本地通信（stdio）；
* 2、基于SSE（Server-Sent Events）的远程通信。是一种基于HTTP协议的单向通信协议，允许服务器以事件流的形式实时向客户端推送数据，而无需客户端明确请求。MCP中的SSE Transport结合了SSE技术和HTTP POST

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/67.png)

* 3、Streamable-http

streamable HTTP是MCP协议在2025年3月引入的一种新传输机制，旨在取代之前的HTTP+SSE传输模式。它的设计理念是在保留SSE优点的同时克服其限制，特别是提供更好的可扩展性和企业环境兼容性。

Streamable HTTP的核心思想是提供一个统一的HTTP端点，同时支持POST和GET方法：

1. POST方法：用于客户端向服务器发送请求和接收响应
2. GET方法（可选）：用于建立SSE流，接收服务器实时推送的消息

与传统HTTP+SSE不同，Streamable HTTP不要求维护单独的初始化连接和消息端点，简化了协议设计并提高了可靠性。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/68.png)

1. **连接**：在streamable HTTP模式下，并没有和SSE模式类似的“连接”过程（在sse\_client调用时），因为无需事先创建SSE连接；
2. **客户端发起初始化请求（Initialize）**。如果是有状态模式，会在返回消息的HTTP头中携带session-id；
3. **客户端发起初始化确认（Initialized）**。此时如果已有session-id（有状态），客户端会首先发起一次HTTP Get请求，以建立独立的SSE通道；
4. **后续正常交互**：普通的交互都是通过Post通道来进行，只有两种情况会使用SSE通道：服务端发起的通知与请求、以及会话恢复的事件发送。

**Streamable HTTP相比SSE的五大优势**

1. 简化的通信模型

传统的HTTP+SSE方法需要两个不同的端点：一个用于建立连接，另一个用于发送消息。而Streamable HTTP提供了一个统一的端点，简化了客户端和服务器之间的交互。

1. 支持无状态模式

Streamable HTTP的一个重要创新是支持完全无状态操作。通过设置sessionIdGenerator: () => undefined，服务器可以在不维护会话状态的情况下处理请求，非常适合无服务器环境。比如：部署在AWS Lambda、Azure Functions等无服务器环境。短暂交互而非长期连接的场景。

1. 更好的可伸缩性

由于Streamable HTTP可以在无状态模式下运行，它非常适合容器化和自动扩展场景。服务器不需要维护长期连接，可以根据请求动态分配资源，显著提高可伸缩性。

这解决了SSE的一个主要问题：当有大量客户端时，每个客户端都需要维持一个长连接，可能导致服务器资源耗尽。使用Streamable HTTP的无状态模式，服务器只在处理请求时分配资源，处理完成后即可释放。

1. 提高的可靠性

Streamable HTTP的简化设计减少了出错机会：

* 会话管理：在有状态模式下，会话ID通过HTTP头而非查询参数传递，减少安全风险
* 重连处理：客户端可以在会话有效期内随时重连，无需复杂的重连逻辑
* 错误恢复：简化的协议使错误处理和恢复更加直观

1. 更好的企业环境兼容性

在企业环境中，代理服务器和防火墙常常会阻止非标准HTTP连接。Streamable HTTP使用标准HTTP通信，大大减少了这类问题：

* 使用标准HTTP POST和GET，无需特殊配置
* 不依赖长连接，减少代理超时问题
* 会话ID通过HTTP头传递，更符合企业安全要求

**5、FastMCP：构建模型上下文协议（MCP）服务器的快速Python方案**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/69.png)

|  |
| --- |
| Python pip install langchain-mcp-adapters |

Java开发MCP服务：[把SpringBoot项目改造成MCP Server非常简单 - 知乎](https://zhuanlan.zhihu.com/p/1892863789579351119)

Python开发MCP服务：FastMCP

开源的MCP服务项目：[MCP servers | Glama](https://glama.ai/mcp/servers)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/70.png)

它提供了一种简单且高效的方法来构建MCP服务器，为开发者提供强大的工具和资源，从而帮助他们为LLMs提供上下文信息。

**FastMCP的主要特性：**

* 快速：高层接口意味着更少的代码和更快的开发速度。
* 简单：构建MCP服务器时，所需的样板代码最少。
* Pythonic：符合Python开发者的直觉，使用起来更自然。
* 完整：致力于提供MCP规格的全面实现（当前某些高级功能仍在开发中）。♂️
* FastMCP正在积极开发中，而MCP规格本身也在不断完善。

**6、MCP的认证和安全**

FastMCP 的 Bearer Token Authentication 认证机制基于 JWT (JSON Web Token) 标准实现，是一种无状态的、基于声明的安全验证体系。

**认证验证流程**

当请求到达服务端时，按以下顺序验证：

* **存在性检查**​：确认请求头包含 Authorization: Bearer <token>
* **结构验证**​：检查JWT格式是否符合规范
* **签名验证**​：使用配置的公钥验证签名有效性
* **声明校验**​：
* iss 必须匹配 BearerAuthProvider 配置的签发方
* aud 必须包含服务端指定的受众
* exp 必须未过期
* **权限检查**​：Token中的 scopes 需满足工具要求的权限

1. 生成 RSA 密钥对

**作用**​：生成用于 JWT 签名和验证的 RSA 密钥对

**详细说明**​：

* 生成一对 RSA 密钥（公钥和私钥）
* 公钥(public\_key)用于验证 Token 签名
* 私钥(private\_key)用于签发 Token

|  |
| --- |
| Python # 1. 生成 RSA 密钥对（生产环境应从安全存储读取） key\_pair = RSAKeyPair.generate() |

1. 配置 Bearer 认证提供方

|  |
| --- |
| Python # 2. 配置认证提供方 auth = BearerAuthProvider(  public\_key=key\_pair.public\_key, # 公钥用于校验签名  issuer="http://localhost", # 令牌签发方标识  audience="my-dev-server", # 目标服务标识 ) |

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/71.png)

1. 生成token

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/72.png)

|  |
| --- |
| Python # 服务器，模拟生成一个token # Generate a token for testing token = key\_pair.create\_token(  subject="dev-user",  issuer="http://localhost",  audience="my-dev-server",  scopes=["read", "write"] # token中包含的数据  expires\_in\_seconds=3600 # 1小时后过期 ) |

1. 客户端Agent传入token
2. 服务器验证和获取Token

get\_access\_token()

* 作用​：从当前请求中提取并解码 JWT
* 返回​：AccessToken 对象，包含以下主要属性：

|  |
| --- |
| Python class AccessToken:  client\_id: str # 对应JWT的'sub'(subject)声明  scopes: List[str] # 对应JWT的'scopes'声明  issuer: str # 对应JWT的'iss'声明  token: str # 对应JWT的Token字符串  expires\_at: int # 对应JWT的'exp'声明 |

**第五章、基于LangGraph的WorkFlow**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/73.png)

工作流是LLM和工具通过预定义的代码路径进行编排的系统。另一方面，智能体是LLM动态指导其自身流程和工具使用的系统，它们都是构建复杂的大模型应用系统的核心组件。

LangGraph的工作流通过有向图（Directed Graph）定义，由节点（Node）、边（Edge）、状态（State）构成，支持条件分支、循环、并行执行。通过**状态图（StateGraph）​**将多个智能体（Agent）和任务节点（Node）组织成可动态调整的流程结构，实现**有状态、可扩展的任务编排**。

LangGraph通过**状态图**将Agent和工作流解耦，让开发者能以“搭积木”的方式构建复杂AI系统：

* **Agent**是“乐高积木”，负责实现具体功能；
* **工作流**是“搭建规则”，定义积木如何组合、何时执行。  
   这种设计既保留了灵活性，又能通过状态管理实现大规模系统的可靠运行。

**一句话总结：Agent 是一种特殊的节点（Node），而工作流是整个图的运行流程（Graph）。Agent 负责“思考+行动”，工作流负责“编排+调度”。**

**1、Graph的相关概念**

**一、StateGraph**

LangGraph 的核心是将Agent，WorkFlow建模为图。您可以使用三个关键组件来定义一个Graph：

1. [State](https://github.langchain.ac.cn/langgraph/concepts/low_level/#state): 表示应用程序当前快照的共享数据结构。它可以是任何 Python 类型，但通常是 TypedDict 或 Pydantic BaseModel。
2. [Nodes](https://github.langchain.ac.cn/langgraph/concepts/low_level/#nodes): 编码代理逻辑的 Python 函数。它们接收当前的 State 作为输入，执行一些计算或副作用，并返回更新后的 State。
3. [Edges](https://github.langchain.ac.cn/langgraph/concepts/low_level/#edges): 根据当前 State 决定接下来执行哪个 Node 的 Python 函数。它们可以是条件分支或固定转换。

通过组合 Nodes 和 Edges，您可以创建复杂的、循环的工作流，这些工作流会随着时间推移演进 State。需要强调的是：Nodes 和 Edges 仅仅是 Python 函数——它们可以包含 LLM 或仅是普通的 Python 代码。

**要构建图，您首先定义**[**状态**](https://github.langchain.ac.cn/langgraph/concepts/low_level/#state)**，然后添加**[**节点**](https://github.langchain.ac.cn/langgraph/concepts/low_level/#nodes)**和**[**边**](https://github.langchain.ac.cn/langgraph/concepts/low_level/#edges)**，最后进行编译。**

|  |
| --- |
| Python graph = graph\_builder.compile(...) |

**二、State状态**

定义图时，您要做的第一件事是定义图的 State。State 由[图的 schema](https://github.langchain.ac.cn/langgraph/concepts/low_level/#schema) 以及指定如何将更新应用到状态的[reducer 函数](https://github.langchain.ac.cn/langgraph/concepts/low_level/#reducers)组成。State 的 schema 将作为图中所有 Nodes 和 Edges 的输入 schema，可以是 TypedDict 或 Pydantic 模型。所有 Nodes 将发出对 State 的更新，然后使用指定的 reducer 函数应用这些更新。

|  |
| --- |
| Python from langgraph.graph.message import add\_messages   class State(TypedDict):  messages: Annotated[list[AnyMessage], add\_messages]  extra\_field: int |

**Reducer函数**

理解节点更新如何应用于 State，Reducer 是关键。State 中的每个键都有其独立的 reducer 函数。如果未明确指定 reducer 函数，则假定对该键的所有更新都应覆盖它。Reducer 有几种不同类型，从默认类型的 reducer 开始：

实例A

|  |
| --- |
| Python from typing\_extensions import TypedDict  class State(TypedDict):  foo: int  bar: list[str] |

在此示例中，未为任何键指定 reducer 函数。假设图的输入是 {"foo": 1, "bar": ["hi"]}。然后假设第一个 Node 返回 {"foo": 2}。这被视为对状态的更新。请注意，Node 不需要返回整个 State schema，而只需返回一个更新。应用此更新后，State 将变为 {"foo": 2, "bar": ["hi"]}。如果第二个节点返回 {"bar": ["bye"]}，则 State 将变为 {"foo": 2, "bar": ["bye"]}。

实例B

|  |
| --- |
| Python from typing import Annotated from typing\_extensions import TypedDict from operator import add  class State(TypedDict):  foo: int  bar: Annotated[list[str], add] |

在此示例中，我们使用 Annotated 类型为第二个键 (bar) 指定了一个 reducer 函数 (operator.add)。请注意，第一个键保持不变。假设图的输入是 {"foo": 1, "bar": ["hi"]}。然后假设第一个 Node 返回 {"foo": 2}。这被视为对状态的更新。请注意，Node 不需要返回整个 State schema，而只需返回一个更新。应用此更新后，State 将变为 {"foo": 2, "bar": ["hi"]}。如果第二个节点返回 {"bar": ["bye"]}，则 State 将变为 {"foo": 2, "bar": ["hi", "bye"]}。请注意，这里的 bar 键是通过将两个列表相加来更新的。

**operator.add 和LangGraph.add\_messages区别**

在许多情况下，将之前的对话历史记录作为消息列表存储在图状态中会很有帮助。为此，我们可以在图状态中添加一个键（通道），用于存储 Message 对象列表，并使用 reducer 函数对其进行注释（参见下面示例中的 messages 键）。reducer 函数对于告诉图如何通过每次状态更新（例如，当节点发送更新时）来更新状态中的 Message 对象列表至关重要。如果您不指定 reducer，每次状态更新都将用最新提供的值覆盖消息列表。如果您只想将消息简单地附加到现有列表中，可以使用 operator.add 作为 reducer。

但是，您可能还需要手动更新图状态中的消息（例如，人机协作）。如果您使用 operator.add，您发送给图的手动状态更新将被附加到现有消息列表，而不是更新现有消息。为避免这种情况，您需要一个能够跟踪消息 ID 并在更新时覆盖现有消息的 reducer。为此，您可以使用预构建的 add\_messages 函数。对于新消息，它只会简单地附加到现有列表，但它也会正确处理现有消息的更新。除了跟踪消息 ID，add\_messages 函数还会在 messages 通道上收到状态更新时尝试将消息反序列化为 LangChain Message 对象。

**MessagesState**

由于在状态中包含消息列表非常常见，因此存在一个预构建的状态 MessagesState，它使得使用消息变得容易。MessagesState 定义了一个单一的 messages 键，它是一个 AnyMessage 对象列表，并使用 add\_messages reducer。通常，需要跟踪的状态不仅仅是消息，因此我们看到人们会对此状态进行子类化并添加更多字段，例如：

|  |
| --- |
| Python from langgraph.graph import MessagesState  class State(MessagesState):  documents: list[str] # 额外增加的key |

**三、节点Node**

在 LangGraph 中，节点通常是 Python 函数（同步或异步），其中**第一个**位置参数是[状态](https://github.langchain.ac.cn/langgraph/concepts/low_level/#state)，（可选地）**第二个**位置参数是“config”，包含可选的[可配置参数](https://github.langchain.ac.cn/langgraph/concepts/low_level/#configuration)（例如 thread\_id）。

|  |
| --- |
| Python from typing\_extensions import TypedDict  from langchain\_core.runnables import RunnableConfig from langgraph.graph import StateGraph  class State(TypedDict):  input: str  results: str  builder = StateGraph(State)   def my\_node(state: State, config: RunnableConfig):  print("In node: ", config["configurable"]["user\_id"])  return {"results": f"Hello, {state['input']}!"}   # The second argument is optional def my\_other\_node(state: State):  return state   builder.add\_node("my\_node", my\_node) builder.add\_node("other\_node", my\_other\_node) ... |

**START 节点**

START 节点是一个特殊节点，它代表将用户输入发送到图的节点。引用此节点的主要目的是确定应首先调用哪些节点。

**END 节点**

END 节点是一个特殊节点，代表一个终止节点。当您想表示哪些边在完成后没有后续操作时，会引用此节点。

**四、边Edge**

边定义了逻辑如何路由以及图如何决定停止。这是代理工作方式以及不同节点之间如何通信的重要组成部分。有几种关键类型的边：

* 普通边：直接从一个节点到下一个节点。
* 条件边：调用函数以确定接下来要转向哪些节点。

一个节点可以有多个出边。如果一个节点有多个出边，则**所有**这些目标节点将在下一个超级步骤中并行执行。

**普通边**

如果您**总是**想从节点 A 转到节点 B，可以直接使用 [add\_edge](https://github.langchain.ac.cn/langgraph/reference/graphs/#langgraph.graph.state.StateGraph.add_edge) 方法。

**条件边**

如果您想**可选地**路由到一条或多条边（或可选地终止），可以使用 [add\_conditional\_edges](https://github.langchain.ac.cn/langgraph/reference/graphs/#langgraph.graph.state.StateGraph.add_conditional_edges) 方法。此方法接受一个节点名称和一个在该节点执行后调用的“路由函数”：

|  |
| --- |
| Python graph.add\_conditional\_edges("node\_a", routing\_function) |

与节点类似，routing\_function 接受图的当前 state 并返回一个值。

默认情况下，routing\_function 的返回值用作下一个要发送状态的节点（或节点列表）的名称。所有这些节点都将在下一个超级步骤中并行运行。

您可以选择提供一个字典，将 routing\_function 的输出映射到下一个节点的名称。

|  |
| --- |
| Python graph.add\_conditional\_edges("node\_a", routing\_function, {True: "node\_b", False: "node\_c"}) |

routing\_function就必须返回：Ture和False。

**2、第一个案例：评估器**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/74.png)

|  |
| --- |
| Python from typing import Literal, TypedDict from pydantic import BaseModel, Field from langgraph.graph import StateGraph, END, START from agent.my\_llm import llm  *# 定义图状态结构（使用TypedDict实现类型安全）* class State(TypedDict):  joke: str *# 生成的冷笑话内容* topic: str *# 用户指定的主题* feedback: str *# 改进建议* funny\_or\_not: str *# 幽默评级*  *# 结构化输出模型（用于LLM评估反馈）* class Feedback(BaseModel):  *"""使用此工具来结构化你的响应"""* grade: Literal["funny", "not funny"] = Field(  description="判断笑话是否幽默",  examples=["funny", "not funny"]  )  feedback: str = Field(  description="若不幽默，提供改进建议",  example="可以加入双关语或意外结局"  )  *# # 配置LLM结构化输出（需提前初始化llm对象）* *# evaluator = llm.with\_structured\_output(Feedback)*  *# 节点定义* def llm\_call\_generator(state: State) -> dict:  *"""LLM生成笑话节点（带反馈优化机制）"""* prompt = (  f"根据反馈改进笑话：{state['feedback']}\n主题：{state['topic']}"  if state.get("feedback")  else f"创作一个关于{state['topic']}的笑话"  )  response = llm.invoke(prompt)  return {"joke": response.content}  def llm\_call\_evaluator(state: State) -> dict:  *"""LLM笑话评估节点（结构化输出）"""* evaluator = llm.bind\_tools([Feedback])  evaluation = evaluator.invoke(  f"评估此笑话的幽默程度：\n{state['joke']}\n"  "注意：幽默应包含意外性或巧妙措辞"  )  evaluation = evaluation.tool\_calls[-1]['args']  return {  "funny\_or\_not": evaluation['grade'],  "feedback": evaluation['feedback']  }  *# 条件路由函数* def route\_joke(state: State) -> str:  *"""动态路由决策（基于评估结果）"""* return (  "Accepted" if state["funny\_or\_not"] == "funny"  else "Rejected + Feedback"  )  *# 构建工作流（使用StateGraph）* optimizer\_builder = StateGraph(State)  *# 添加节点（可替换为MCP工具[5](@ref)）* optimizer\_builder.add\_node("generator", llm\_call\_generator) *# 对应@mcp.tool()* optimizer\_builder.add\_node("evaluator", llm\_call\_evaluator)  *# 构建连接关系* optimizer\_builder.add\_edge(START, "generator") optimizer\_builder.add\_edge("generator", "evaluator") optimizer\_builder.add\_conditional\_edges(  "evaluator",  route\_joke,  {  "Accepted": END, *# 合格则结束* "Rejected + Feedback": "generator" *# 不合格则循环优化* } )  *# 编译工作流（可导出为MCP服务[6](@ref)）* optimizer\_workflow = optimizer\_builder.compile()  *# 执行示例（主题可动态替换）* *# result = optimizer\_workflow.invoke({"topic": "程序员"})* *# print(f"最终笑话：{result['joke']}")* |

**3、综合案例： 记忆存储 + 人工干预 + 工具节点**

需求： 开发一个智能小秘书。要求：可以帮我查询火车票，帮我生成分析图表，可以帮我获取网络上的数据。

**一、MCP工具**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/75.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/76.png)

**二、自定义工具调用**

调用工具： 异步，并发，参数验证，异常处理，人工介入等各种功能和要求。

|  |
| --- |
| Python class BasicToolNode:  *"""异步工具节点，用于并发执行AIMessage中请求的工具调用*   *功能：*  *1. 接收工具列表并建立名称索引*  *2. 并发执行消息中的工具调用请求*  *3. 自动处理同步/异步工具适配*  *"""*  def \_\_init\_\_(self, tools: list) -> None:  *"""初始化工具节点*  *Args:*  *tools: 工具列表，每个工具需包含name属性*  *"""* self.tools\_by\_name = {tool.name: tool for tool in tools} *# 建立工具名称索引*  async def \_\_call\_\_(self, inputs: Dict[str, Any]) -> Dict[str, List[ToolMessage]]:  *"""异步调用入口*  *Args:*  *inputs: 输入字典，需包含"messages"字段*  *Returns:*  *包含ToolMessage列表的字典*  *Raises:*  *ValueError: 当输入无效时抛出*  *"""*  *# 1. 输入验证* if not (messages := inputs.get("messages")):  raise ValueError("输入数据中未找到消息内容") *# 改进后的中文错误提示* message = messages[-1] *# 取最新消息*   *# 2. 并发执行工具调用* outputs = await self.\_execute\_tool\_calls(message.tool\_calls)  return {"messages": outputs}   async def \_execute\_tool\_calls(self, tool\_calls: List[Dict]) -> List[ToolMessage]:  *"""执行实际工具调用*  *Args:*  *tool\_calls: 工具调用请求列表*  *Returns:*  *ToolMessage结果列表*  *"""*  async def \_invoke\_tool(tool\_call: Dict) -> ToolMessage:  *"""执行单个工具调用*  *Args:*  *tool\_call: 工具调用请求字典，需包含name/args/id字段*  *Returns:*  *封装的ToolMessage*  *Raises:*  *KeyError: 工具未注册时抛出*  *RuntimeError: 工具调用失败时抛出*  *"""* try:  *# 3. 异步调用工具* tool = self.tools\_by\_name.get(tool\_call["name"])  if not tool:  raise KeyError(f"未注册的工具: {tool\_call['name']}")   if hasattr(tool, 'ainvoke'): *# 优先使用异步方法* tool\_result = await tool.ainvoke(tool\_call["args"])  else: *# 同步工具通过线程池转异步* loop = asyncio.get\_running\_loop()  tool\_result = await loop.run\_in\_executor(  None, *# 使用默认线程池* tool.invoke, *# 同步调用方法* tool\_call["args"] *# 参数* )   *# 4. 构造ToolMessage* return ToolMessage(  content=json.dumps(tool\_result, ensure\_ascii=False),  name=tool\_call["name"],  tool\_call\_id=tool\_call["id"],  )  except Exception as e:  raise RuntimeError(f"工具调用失败: {tool\_call['name']}") from e   *# 5. 并发执行所有工具调用*  *# '''*  *# asyncio.gather() 是 Python 异步编程中用于并发调度多个协程的核心函数，其核心行为包括：*  *# 并发执行：所有传入的协程会被同时调度到事件循环中，通过非阻塞 I/O 实现并行处理。*  *# 结果收集：按输入顺序返回所有协程的结果（或异常），与任务完成顺序无关。*  *# 异常处理：默认情况下，任一任务失败会立即取消其他任务并抛出异常；若设置 return\_exceptions=True，则异常会作为结果返回。*  *# '''* try:  return await asyncio.gather(\*[  \_invoke\_tool(tool\_call) for tool\_call in tool\_calls  ])  except Exception as e:  raise RuntimeError("并发执行工具时发生错误") from e |

**三、LangGraph自带的工具执行**

**ToolNode**

一个运行在最后一条 AIMessage 中调用的工具的节点。

它可以在 StateGraph 中与“messages”状态键一起使用（或通过 ToolNode 的“messages\_key”传递自定义键）。如果请求多个工具调用，它们将并行运行。输出将是一个 ToolMessages 列表，每个工具调用对应一个。

**对比**

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/77.png)

**点击图片可查看完整电子表格**

|  |
| --- |
| Python from langgraph.prebuilt import ToolNode from langchain\_core.tools import tool  @tool def calculator(a: int, b: int) -> int:  """Add two numbers."""  return a + b  tool\_node = ToolNode([calculator]) |

**四、人机协作 (Human-in-the-loop)**

要在代理或工作流中审查、编辑和批准工具调用，请使用LangGraph的人机协作功能，以在工作流的任何点启用人工干预。这在大型语言模型 (LLM) 驱动的应用程序中特别有用，因为模型输出可能需要验证、修正或额外的上下文。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/78.png)

您可以使用interrupt和Command实现四种典型的应用场景

* **批准或拒绝：**在关键步骤（例如API调用）之前暂停图，以审查并批准操作。如果操作被拒绝，您可以阻止图执行该步骤，并可能采取替代操作。这种模式通常涉及根据人工输入路由图。
* **编辑图状态：**暂停图以审查和编辑图状态。这对于纠正错误或用附加信息更新状态很有用。这种模式通常涉及用人工输入更新状态。
* **审查工具调用：**在工具执行之前，暂停图以审查和编辑LLM请求的工具调用。
* **验证人工输入：**在进行下一步之前，暂停图以验证人工输入。

这些功能功能的主要接口是[interrupt](https://github.langchain.ac.cn/langgraph/how-tos/human_in_the_loop/add-human-in-the-loop/)函数。在节点内调用interrupt将暂停执行。通过传入[Command](https://github.langchain.ac.cn/langgraph/concepts/low_level/#command)，可以恢复执行并接收来自人工的新输入。interrupt在功能上类似于 Python 的内置input()。

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/79.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/80.png)

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/81.png)

**使用 interrupt 暂停(中断)**

要在图中使用 interrupt，您需要

1. **指定一个检查点器**，用于在每一步之后保存图状态。
2. 在适当位置**调用 interrupt()，第二种：传入interrupt\_before[节点名字]**。
3. 使用**线程 ID** **运行图**，直到触发 interrupt。
4. 使用 invoke/ainvoke/stream/astream
5. **恢复执行： interrupt\_before--->agent.astream(None, config, stream\_mode='values')**
6. **恢复执行：interrupt()---->agent.astream(Command(resume={"answer": user\_input}), config, stream\_mode='values')**

**4、综合案例： 多智能体的监督和管理**

**该课程在：携程AI智能助手项目增加了，链接：https://www.mashibing.com/subject/167?activeNav=1&courseNo=2801&courseVersionId=3750**

[**langgraph-supervisor**](https://github.langchain.ac.cn/langgraph/agents/multi-agent/#supervisor) **可以为多个Agent的工作流构建一个监督者。**

**监督者**  是一种多代理架构，其中**专业化**代理由一个中心**监督代理**协调。监督代理控制所有通信流和任务委托，根据当前上下文和任务要求决定调用哪个代理

![](新版Langchain+LangGraph+MCP的智能体和工作流开发(最终版)_media/82.png)

多智能体交互中的一个常见模式是**交接**，即一个智能体将控制权*移交给*另一个智能体。交接允许您指定：

* **目的地**：要导航到的目标智能体（例如，要前往的节点名称）
* **载荷**：[要传递给该智能体的信息](https://github.langchain.ac.cn/langgraph/concepts/multi_agent/#communication-and-state-management)（例如，状态更新）

要在LangGraph中实现交接，智能体节点可以返回[Command](https://github.langchain.ac.cn/langgraph/concepts/low_level/#command)对象，该对象允许您结合控制流和状态更新。一种常见模式是将交接封装在工具调用中，例如：

|  |
| --- |
| Python @tool *# 使用tool装饰器将该函数注册为可调用工具* def transfer\_to\_bob():  """交接给bob"""   *# 工具功能描述：执行向bob交接的操作*  return Command(  *# 指定要跳转到的目标代理(节点)名称*  goto="bob",  *# 要传递给目标代理的数据内容*  update={"my\_state\_key": "my\_state\_value"},  *# 指示LangGraph需要导航到父级图中的代理节点*  graph=Command.PARENT,  ) |