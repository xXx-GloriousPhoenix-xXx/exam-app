// npm install docx
// node generate_exam_doc.js

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageBreak, BorderStyle
  } = require('docx');
  const fs = require('fs');
  
  // ─── helpers ────────────────────────────────────────────────────────────────
  
  function h1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text, bold: true, size: 32, font: 'Arial' })],
      spacing: { before: 400, after: 200 },
    });
  }
  
  function h2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, size: 28, font: 'Arial' })],
      spacing: { before: 300, after: 160 },
    });
  }
  
  function para(text, opts = {}) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 100, after: 100, line: 276 },
      children: [new TextRun({ text, font: 'Arial', size: 24, ...opts })],
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, font: 'Arial', size: 24 })],
    });
  }
  
  function code(text) {
    return new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [new TextRun({
        text,
        font: 'Courier New',
        size: 20,
      })],
      indent: { left: 720 },
    });
  }
  
  function pageBreak() {
    return new Paragraph({ children: [new PageBreak()] });
  }
  
  function divider() {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 1 } },
      spacing: { before: 200, after: 200 },
      children: [],
    });
  }
  
  // ─── question sections ───────────────────────────────────────────────────────
  
  const sections = [];
  
  // ── Q1 ──────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 1. Поняття паралельних, псевдопаралельних та розподілених обчислень'),
    divider(),
  
    h2('1. Послідовні обчислення (Sequential Computing)'),
    para('Традиційна модель виконання програм, за якої інструкції виконуються одна за одною на одному процесорі. Кожна наступна інструкція виконується лише після завершення попередньої. Це найпростіша модель, але вона не дозволяє ефективно використовувати сучасні багатоядерні процесори.'),
  
    h2('2. Псевдопаралельні обчислення (Pseudo-Parallel / Concurrency)'),
    para('Псевдопаралельність — це ілюзія одночасного виконання кількох задач на одному процесорі шляхом швидкого перемикання між ними. Операційна система розподіляє процесорний час між потоками/процесами за допомогою планувальника (scheduler).'),
    para('Ключові характеристики:'),
    bullet('Одночасно виконується лише одна задача фізично'),
    bullet('Переключення відбувається так швидко, що користувач сприймає їх як паралельні'),
    bullet('Залежить від кількості квантів часу (time slices), виділених планувальником'),
    bullet('Типовий приклад — запуск кількох програм на однопроцесорній машині'),
    para('У Java псевдопаралельність виникає при запуску кількох потоків (Thread) на однопроцесорній машині або коли потоків більше, ніж фізичних ядер. JVM + ОС вирішують, коли перемикати потоки.'),
  
    h2('3. Справжні паралельні обчислення (True Parallelism)'),
    para('Паралельні обчислення — одночасне виконання кількох обчислень на різних процесорах або ядрах у один і той же момент часу. Потребує багатоядерного процесора або кількох процесорів.'),
    para('Основні цілі паралельних обчислень:'),
    bullet('Прискорення виконання задачі (speed-up)'),
    bullet('Збільшення масштабованості системи'),
    bullet('Вирішення задач, що неможливі на одному процесорі через обмеження пам\'яті чи часу'),
    para('Приклад у Java: використання ForkJoinPool або паралельних потоків (parallelStream()) на багатоядерній машині — різні гілки обчислень реально виконуються одночасно на різних ядрах.'),
    code('// Java parallel stream — справжній паралелізм на багатоядерному CPU'),
    code('List<Integer> data = IntStream.range(0, 1_000_000).boxed().collect(Collectors.toList());'),
    code('long sum = data.parallelStream().mapToLong(Integer::longValue).sum();'),
  
    h2('4. Розподілені обчислення (Distributed Computing)'),
    para('Розподілені обчислення — це модель, за якої задача вирішується групою комп\'ютерів, з\'єднаних мережею. Кожен вузол (node) має власну пам\'ять і процесор; вони взаємодіють через обмін повідомленнями (message passing).'),
    para('Відмінності від паралельних обчислень:'),
    bullet('Паралельні: спільна пам\'ять (shared memory), швидка взаємодія через RAM'),
    bullet('Розподілені: розподілена пам\'ять (distributed memory), взаємодія через мережу (MPI, HTTP, gRPC)'),
    bullet('Розподілені системи стійкіші до відмов (fault tolerance)'),
    bullet('Розподілені можуть масштабуватись на тисячі вузлів (HPC-кластери, хмарні платформи)'),
    para('Технології: MPI (Message Passing Interface) — стандарт для наукових обчислень на кластерах; Apache Spark, Hadoop — для великих даних; мікросервіси на Spring Boot — для бізнес-задач.'),
  
    h2('5. Порівняльна таблиця'),
    bullet('Послідовні: 1 процесор, 1 задача одночасно, проста модель'),
    bullet('Псевдопаралельні: 1 процесор, багато задач (квантування часу), ілюзія паралелізму'),
    bullet('Паралельні: N ядер/CPU, N задач одночасно, спільна пам\'ять'),
    bullet('Розподілені: N вузлів у мережі, повідомлення між процесами, distributed memory'),
  
    pageBreak(),
  ]);
  
  // ── Q3 ──────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 3. Способи розпаралелювання'),
    divider(),
  
    h2('1. Що таке розпаралелювання'),
    para('Розпаралелювання (parallelization) — процес перетворення послідовного алгоритму або програми в паралельну, щоб кілька процесорів або потоків могли виконувати роботу одночасно. Мета — зменшення часу виконання.'),
  
    h2('2. Декомпозиція задач (Task Decomposition)'),
    para('Задача розбивається на незалежні підзадачі, кожна з яких може виконуватись окремим потоком або процесом. Підходить, коли підзадачі логічно відокремлені.'),
    para('Приклад: у системі масового обслуговування кожен клієнт обробляється незалежним потоком із FixedThreadPool:'),
    code('ExecutorService pool = Executors.newFixedThreadPool(4);'),
    code('for (Client c : clients) pool.submit(() -> processClient(c));'),
  
    h2('3. Декомпозиція даних (Data Decomposition)'),
    para('Дані розбиваються на частини (partition), і кожна частина обробляється окремим потоком або процесом. Найбільш поширений підхід для числових обчислень.'),
    para('Варіанти розбиття:'),
    bullet('Рядкове (стрічкове) — кожен потік отримує набір рядків матриці (стрічковий алгоритм множення матриць)'),
    bullet('Стовпчикове — розбиття по стовпцях'),
    bullet('Блокове — матриця розбивається на блоки (алгоритм Фокса/Кеннона)'),
    bullet('Циклічне (round-robin) — елементи розподіляються по колу між потоками'),
    code('// Стрічкова декомпозиція: кожен потік множить свою смугу рядків'),
    code('int rowsPerThread = N / numThreads;'),
    code('for (int t = 0; t < numThreads; t++) {'),
    code('  final int start = t * rowsPerThread;'),
    code('  final int end = (t == numThreads-1) ? N : start + rowsPerThread;'),
    code('  pool.submit(() -> multiplyRows(A, B, C, start, end));'),
    code('}'),
  
    h2('4. Розпаралелювання рекурсії — ForkJoin'),
    para('Патерн "розділяй та володарюй" (divide and conquer): задача рекурсивно розбивається на підзадачі до досягнення базового випадку, підзадачі вирішуються паралельно, результати об\'єднуються.'),
    code('class SumTask extends RecursiveTask<Long> {'),
    code('  protected Long compute() {'),
    code('    if (size <= THRESHOLD) return sequentialSum();'),
    code('    SumTask left = new SumTask(array, lo, mid);'),
    code('    SumTask right = new SumTask(array, mid, hi);'),
    code('    left.fork(); // асинхронно'),
    code('    return right.compute() + left.join();'),
    code('  }'),
    code('}'),
  
    h2('5. Конвеєрне розпаралелювання (Pipeline)'),
    para('Обробка організована як конвеєр: кожен потік виконує свій етап, результат передається наступному. Аналог заводського конвеєра. Реалізується через чергу (BlockingQueue) між потоками (Producer-Consumer).'),
    bullet('Потік 1: зчитує дані з файлу → queue1'),
    bullet('Потік 2: парсить дані з queue1 → queue2'),
    bullet('Потік 3: записує результат з queue2'),
  
    h2('6. Колективне розпаралелювання (MPI)'),
    para('У розподілених системах використовуються колективні операції MPI, де всі процеси беруть участь одночасно: MPI_Bcast (broadcast), MPI_Scatter, MPI_Gather, MPI_Reduce, MPI_Allreduce. Це дозволяє ефективно координувати велику кількість вузлів.'),
  
    h2('7. Критерії вибору способу'),
    bullet('Незалежні задачі → Task Decomposition (FixedThreadPool)'),
    bullet('Великі масиви/матриці → Data Decomposition (стрічковий алгоритм)'),
    bullet('Рекурсивні задачі → ForkJoin'),
    bullet('Потоки даних → Pipeline (Producer-Consumer)'),
    bullet('Кластер вузлів → MPI-колективи'),
  
    pageBreak(),
  ]);
  
  // ── Q5 ──────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 5. Способи підвищення продуктивності комп\'ютерів. Суперкомп\'ютери'),
    divider(),
  
    h2('1. Апаратні способи підвищення продуктивності'),
    para('Збільшення тактової частоти (Clock Speed): до 2000-х років процесори ставали швидшими просто завдяки зростанню частоти. Але фізичні обмеження (теплоємність, споживання енергії) призупинили цей шлях приблизно на 3–4 ГГц.'),
    para('Багатоядерні процесори (Multi-core): замість одного швидшого ядра — кілька ядер на одному чіпі. Сучасні CPU мають від 4 до 128 ядер. Дозволяє справжній паралелізм, але вимагає паралельного програмування.'),
    bullet('Intel Core i9: до 24 ядер'),
    bullet('AMD EPYC: до 128 ядер'),
    bullet('ARM Neoverse: серверні чіпи для хмарних платформ'),
    para('Векторні інструкції (SIMD — Single Instruction Multiple Data): одна інструкція обробляє кілька елементів даних одночасно. SSE, AVX, AVX-512 у x86. Використовується для мультимедіа, ML, наукових обчислень.'),
    para('GPU (Graphics Processing Unit): тисячі простих ядер для масово-паралельних обчислень. NVIDIA CUDA, AMD ROCm. Використовується у ML/AI, симуляціях, рендерингу.'),
    para('Кеш-пам\'ять (Cache Hierarchy): L1, L2, L3 кеш зменшують затримку доступу до даних. Алгоритми, що враховують кеш (cache-friendly), суттєво швидші.'),
  
    h2('2. Програмні способи підвищення продуктивності'),
    bullet('Паралельне програмування (потоки, MPI, CUDA)'),
    bullet('Оптимізація алгоритмів (краща складність O(n))'),
    bullet('Компіляторна оптимізація (JIT у JVM, -O2/-O3 у GCC)'),
    bullet('Асинхронне програмування (non-blocking I/O)'),
    bullet('Профілювання та усунення вузьких місць (bottleneck)'),
  
    h2('3. Суперкомп\'ютери'),
    para('Суперкомп\'ютер — обчислювальна система з найвищою на момент створення продуктивністю. Використовується для задач, що потребують величезних обчислень: кліматичне моделювання, ядерні симуляції, геноміка, AI-тренування.'),
    para('Архітектура суперкомп\'ютера:'),
    bullet('Тисячі або мільйони процесорних ядер'),
    bullet('Надшвидкі мережі (InfiniBand зі швидкістю до 400 Гбіт/с)'),
    bullet('Паралельні файлові системи (Lustre, GPFS)'),
    bullet('Прискорювачі: GPU (NVIDIA), TPU (Google), FPGA'),
    para('Класифікація за TOP500 (рейтинг суперкомп\'ютерів): продуктивність вимірюється у FLOPS (FLoating Point Operations Per Second):'),
    bullet('TFLOPS (10¹²) — терафлопс'),
    bullet('PFLOPS (10¹⁵) — петафлопс'),
    bullet('EFLOPS (10¹⁸) — ексафлопс'),
    para('Станом на 2024 рік, Frontier (Oak Ridge, США) — перший ексафлопсний суперкомп\'ютер (~1.2 EFLOPS). Aurora (Argonne, США) — другий ексафлопс. El Capitan — претендує на 2+ EFLOPS.'),
    para('Програмне забезпечення суперкомп\'ютерів: MPI для міжпроцесної взаємодії, OpenMP для багатопоточності в межах вузла, CUDA/ROCm для GPU, BLAS/LAPACK для лінійної алгебри, Slurm/PBS для планування задач.'),
  
    h2('4. Кластери та грід-системи'),
    para('Кластер — група стандартних серверів (commodity hardware), з\'єднаних мережею. Дешевше за суперкомп\'ютер, але менш продуктивний. Використовується в університетах, дата-центрах.'),
    para('Грід (Grid) — географічно розподілена інфраструктура, що об\'єднує ресурси різних організацій через Інтернет. Приклад: WLCG (Worldwide LHC Computing Grid) для обробки даних Великого адронного колайдера.'),
  
    pageBreak(),
  ]);
  
  // ── Q7 ──────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 7. Поняття потоку обчислень. Створення та запуск потоків в Java'),
    divider(),
  
    h2('1. Що таке потік (Thread)'),
    para('Потік (thread) — це найменша одиниця виконання, керована операційною системою. Потоки існують всередині процесу і поділяють з ним адресний простір (heap, статичні змінні), але мають власний стек викликів (call stack), лічильник команд (program counter) та регістри.'),
    para('Процес vs Потік:'),
    bullet('Процес: незалежний адресний простір, важке перемикання контексту, IPC для зв\'язку'),
    bullet('Потік: спільна пам\'ять процесу, легке перемикання, пряма взаємодія через змінні'),
    para('У JVM кожна програма має щонайменше два потоки: main (головний потік) та GC (збирач сміття). При створенні нового Thread JVM запитує ОС створити нативний потік.'),
  
    h2('2. Спосіб 1: Наслідування класу Thread'),
    code('class MyThread extends Thread {'),
    code('    @Override'),
    code('    public void run() {'),
    code('        System.out.println("Thread: " + Thread.currentThread().getName());'),
    code('        for (int i = 0; i < 5; i++) {'),
    code('            System.out.println(i);'),
    code('        }'),
    code('    }'),
    code('}'),
    code('// Запуск:'),
    code('MyThread t = new MyThread();'),
    code('t.start(); // НЕ t.run()! run() — виклик методу, start() — новий потік'),
    para('Недолік: Java не підтримує множинне наслідування, тому якщо клас вже наслідує інший, цей спосіб неможливий.'),
  
    h2('3. Спосіб 2: Реалізація інтерфейсу Runnable'),
    code('class MyTask implements Runnable {'),
    code('    @Override'),
    code('    public void run() {'),
    code('        System.out.println("Task in: " + Thread.currentThread().getName());'),
    code('    }'),
    code('}'),
    code('// Запуск:'),
    code('Thread t = new Thread(new MyTask());'),
    code('t.start();'),
    code('// Або з lambda (Java 8+):'),
    code('Thread t2 = new Thread(() -> System.out.println("Lambda thread"));'),
    code('t2.start();'),
    para('Перевага: клас може реалізовувати кілька інтерфейсів, задача відокремлена від механізму запуску.'),
  
    h2('4. Спосіб 3: Callable та Future (з результатом)'),
    code('Callable<Integer> task = () -> {'),
    code('    int sum = 0;'),
    code('    for (int i = 1; i <= 100; i++) sum += i;'),
    code('    return sum;'),
    code('};'),
    code('ExecutorService executor = Executors.newSingleThreadExecutor();'),
    code('Future<Integer> future = executor.submit(task);'),
    code('System.out.println("Result: " + future.get()); // блокує до завершення'),
    code('executor.shutdown();'),
    para('На відміну від Runnable, Callable<V> повертає значення і може кидати checked exceptions.'),
  
    h2('5. Метод start() vs run()'),
    para('start() — просить JVM створити новий нативний потік і викликати run() у ньому. run() — звичайний метод, викликається у поточному потоці без паралелізму. Помилка виклику run() замість start() — типова помилка початківців, яка не призводить до паралельного виконання.'),
  
    h2('6. Стани потоку (Thread State)'),
    bullet('NEW — потік створено, start() ще не викликано'),
    bullet('RUNNABLE — виконується або готовий до виконання'),
    bullet('BLOCKED — чекає на монітор (synchronized)'),
    bullet('WAITING — чекає безстроково (wait(), join())'),
    bullet('TIMED_WAITING — чекає обмежений час (sleep(n), wait(n))'),
    bullet('TERMINATED — завершив виконання'),
  
    pageBreak(),
  ]);
  
  // ── Q9 ──────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 9. Клас Thread'),
    divider(),
  
    h2('1. Ієрархія та основні поля'),
    para('java.lang.Thread реалізує Runnable. Кожен потік має: name (ім\'я), priority (1–10), daemon (демон чи ні), id (унікальний long), group (ThreadGroup).'),
  
    h2('2. Конструктори'),
    code('new Thread()                          // run() порожній'),
    code('new Thread(Runnable target)           // виконує target.run()'),
    code('new Thread(Runnable target, String name)'),
    code('new Thread(ThreadGroup group, Runnable target, String name)'),
    code('new Thread(ThreadGroup group, Runnable target, String name, long stackSize)'),
  
    h2('3. Основні методи'),
    para('Методи екземпляра:'),
    bullet('start() — запускає потік; може бути викликаний лише один раз'),
    bullet('run() — тіло потоку; перевизначається або задається через Runnable'),
    bullet('join() — очікує завершення потоку; join(ms) — з таймаутом'),
    bullet('interrupt() — встановлює прапор переривання; сам по собі не зупиняє'),
    bullet('isAlive() — true, якщо потік запущено і ще не завершено'),
    bullet('getName() / setName(String) — ім\'я потоку'),
    bullet('getPriority() / setPriority(int) — пріоритет 1 (MIN) .. 10 (MAX), за замовч. 5 (NORM)'),
    bullet('isDaemon() / setDaemon(boolean) — демон-потік завершується разом з main'),
    bullet('getState() — повертає Thread.State'),
    para('Статичні методи:'),
    bullet('Thread.currentThread() — поточний потік'),
    bullet('Thread.sleep(ms) — пауза без звільнення локів, кидає InterruptedException'),
    bullet('Thread.yield() — підказка планувальнику поступитися ядром'),
    bullet('Thread.interrupted() — перевіряє та скидає прапор переривання'),
  
    h2('4. Пріоритети потоків'),
    para('Пріоритет — підказка планувальнику ОС; не гарантія. Потік з вищим пріоритетом має більше шансів отримати CPU, але поведінка залежить від ОС і JVM-реалізації.'),
    code('Thread t = new Thread(() -> { /* ... */ });'),
    code('t.setPriority(Thread.MAX_PRIORITY); // 10'),
    code('t.start();'),
    para('Константи: Thread.MIN_PRIORITY = 1, Thread.NORM_PRIORITY = 5, Thread.MAX_PRIORITY = 10.'),
  
    h2('5. Демон-потоки'),
    para('Daemon thread — фоновий потік обслуговування (GC, таймери). JVM завершується, коли всі не-демон-потоки закінчились, не чекаючи демонів.'),
    code('Thread daemon = new Thread(() -> { while(true) { /* фоновий моніторинг */ } });'),
    code('daemon.setDaemon(true); // ОБОВ\'ЯЗКОВО до start()'),
    code('daemon.start();'),
  
    h2('6. Переривання (Interruption)'),
    para('interrupt() не вбиває потік, а встановлює прапор isInterrupted(). Методи, що кидають InterruptedException (sleep, wait, join), реагують на нього негайно. Потік сам відповідає за реакцію:'),
    code('Thread t = new Thread(() -> {'),
    code('    while (!Thread.currentThread().isInterrupted()) {'),
    code('        try {'),
    code('            Thread.sleep(1000);'),
    code('            doWork();'),
    code('        } catch (InterruptedException e) {'),
    code('            Thread.currentThread().interrupt(); // відновити прапор'),
    code('            break;'),
    code('        }'),
    code('    }'),
    code('});'),
  
    h2('7. Групи потоків (ThreadGroup)'),
    para('ThreadGroup об\'єднує потоки в ієрархічні групи для масового управління (interrupt всієї групи). Але в сучасному Java замість ThreadGroup рекомендують ExecutorService.'),
  
    pageBreak(),
  ]);
  
  // ── Q11 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 11. Стрічковий алгоритм паралельного множення матриць'),
    divider(),
  
    h2('1. Задача'),
    para('Множення матриць A (N×M) на B (M×K) з отриманням C (N×K). Кожен елемент C[i][j] = ∑ A[i][k]*B[k][j]. Складність послідовного алгоритму — O(N³). Для великих матриць — основний кандидат для розпаралелення.'),
  
    h2('2. Ідея стрічкового алгоритму'),
    para('Матриця A розрізається на горизонтальні "стрічки" (smуги рядків), кожна стрічка призначається окремому потоку. Матриця B передається всім потокам цілком (read-only, без конфліктів). Кожен потік обчислює свою частину матриці C незалежно від інших — синхронізація не потрібна під час обчислень.'),
  
    h2('3. Реалізація в Java'),
    code('public class StripedMatrixMultiply {'),
    code('    static double[][] A, B, C;'),
    code('    static int N, M, K;'),
    code(''),
    code('    static class Worker extends Thread {'),
    code('        int startRow, endRow;'),
    code('        Worker(int s, int e) { startRow = s; endRow = e; }'),
    code(''),
    code('        @Override public void run() {'),
    code('            for (int i = startRow; i < endRow; i++)'),
    code('                for (int j = 0; j < K; j++) {'),
    code('                    C[i][j] = 0;'),
    code('                    for (int k = 0; k < M; k++)'),
    code('                        C[i][j] += A[i][k] * B[k][j];'),
    code('                }'),
    code('        }'),
    code('    }'),
    code(''),
    code('    public static void multiply(int numThreads) throws InterruptedException {'),
    code('        int rowsPerThread = N / numThreads;'),
    code('        Thread[] threads = new Thread[numThreads];'),
    code('        for (int t = 0; t < numThreads; t++) {'),
    code('            int start = t * rowsPerThread;'),
    code('            int end = (t == numThreads - 1) ? N : start + rowsPerThread;'),
    code('            threads[t] = new Worker(start, end);'),
    code('            threads[t].start();'),
    code('        }'),
    code('        for (Thread th : threads) th.join();'),
    code('    }'),
    code('}'),
  
    h2('4. Аналіз продуктивності'),
    para('Теоретичне прискорення: S(p) = T₁ / Tₚ ≈ p (лінійне). На практиці обмежено: накладними витратами на створення потоків, балансуванням навантаження (якщо N не ділиться рівно), доступом до B (кеш-промахи при стовпчиковому читанні).'),
    para('Оптимізація: транспонувати B перед множенням — тоді доступ до B[j][k] стає рядковим і кеш-ефективним:'),
    code('// Транспонуємо B один раз:'),
    code('double[][] BT = new double[K][M];'),
    code('for (int i = 0; i < M; i++) for (int j = 0; j < K; j++) BT[j][i] = B[i][j];'),
    code('// Тоді: C[i][j] += A[i][k] * BT[j][k]; — обидва масиви читаються рядково'),
  
    h2('5. Порівняння з блоковим алгоритмом'),
    bullet('Стрічковий: простий, легко реалізувати, добре для shared-memory'),
    bullet('Блоковий (Фокс, Кеннон): краща кеш-ефективність, підходить для distributed memory (MPI)'),
    bullet('Стрічковий + транспозиція ≈ блоковий за кеш-поведінкою на практиці'),
  
    pageBreak(),
  ]);
  
  // ── Q13 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 13. Алгоритм Кеннона паралельного множення матриць'),
    divider(),
  
    h2('1. Передумови та ідея'),
    para('Алгоритм Кеннона (1969) — блоковий алгоритм для множення матриць у розподіленій пам\'яті на решітці процесів p×p. Оптимізує комунікаційну складність. Матриці A і B розбиваються на блоки p×p підматриць, розміщених на решітці процесів.'),
    para('Ключова ідея: замість широкомовного розсилання (broadcast), кожен процес у кожному кроці отримує лише один блок від сусіда (shift). Це зменшує обсяг комунікацій.'),
  
    h2('2. Кроки алгоритму'),
    bullet('Крок 0: початкове зміщення — рядок i матриці A зсувається на i позицій вліво; стовпець j матриці B зсувається на j позицій вгору'),
    bullet('Крок 1..p: кожен процес (i,j) множить свій поточний блок A_ij на B_ij і додає до C_ij'),
    bullet('Після кожного кроку: блоки A зсуваються на 1 вліво (циклічно по рядку), блоки B зсуваються на 1 вгору (циклічно по стовпцю)'),
    bullet('Після p кроків: C_ij = ∑ A_ik * B_kj — повне множення'),
  
    h2('3. Складність'),
    bullet('Обчислення: O(N³/p) на кожен процес'),
    bullet('Комунікації: O(N²/p * √p) = O(N²/√p) — краще ніж наївний broadcast O(N²)'),
    bullet('Загальна кількість кроків: √p'),
  
    h2('4. Реалізація (концептуальна схема в Java / MPI-псевдокод)'),
    code('// Ініціалізація: визначення позиції процесу у решітці'),
    code('int gridSize = (int) Math.sqrt(numProcesses);'),
    code('int myRow = rank / gridSize;'),
    code('int myCol = rank % gridSize;'),
    code(''),
    code('// Початкове зміщення A і B (через MPI_Sendrecv)'),
    code('// shiftA(myRow) — зсунути блок A на myRow позицій вліво'),
    code('// shiftB(myCol) — зсунути блок B на myCol позицій вгору'),
    code(''),
    code('// p кроків:'),
    code('for (int step = 0; step < gridSize; step++) {'),
    code('    localC += localA * localB; // локальне множення'),
    code('    // зсунути A на 1 вліво, B на 1 вгору'),
    code('    MPI_Sendrecv(localA, leftNeighbor, ..., localA, rightNeighbor, ...);'),
    code('    MPI_Sendrecv(localB, topNeighbor, ...,  localB, bottomNeighbor, ...);'),
    code('}'),
  
    h2('5. Відмінності від алгоритму Фокса'),
    bullet('Кеннон: зсуви (shift), O(N²/√p) комунікацій, потребує квадратну решітку'),
    bullet('Фокс: broadcast по рядку + зсув B по стовпцю, більший трафік, але простіша реалізація'),
    bullet('Обидва: O(N³/p) обчислень, ефективні для великих матриць на кластерах'),
  
    pageBreak(),
  ]);
  
  // ── Q15 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 15. Управління потоками в Java'),
    divider(),
  
    h2('1. Огляд засобів управління'),
    para('Java надає кілька рівнів управління потоками: низький рівень (Thread API), середній (synchronized, wait/notify), високий (java.util.concurrent — Lock, Semaphore, CountDownLatch, ExecutorService).'),
  
    h2('2. Запуск та очікування'),
    code('Thread t = new Thread(() -> doWork());'),
    code('t.start();         // запуск'),
    code('t.join();          // очікування завершення'),
    code('t.join(5000);      // очікування максимум 5 секунд'),
    code('boolean alive = t.isAlive(); // перевірка'),
  
    h2('3. Пауза та поступка'),
    code('Thread.sleep(1000);  // пауза 1 секунда (InterruptedException!)'),
    code('Thread.yield();      // підказка ОС поступитися CPU'),
  
    h2('4. Переривання (Interrupt)'),
    para('Механізм кооперативного зупинення потоку. interrupt() встановлює прапор; потік сам перевіряє і реагує:'),
    code('// Відправник:'),
    code('worker.interrupt();'),
    code('// Отримувач:'),
    code('while (!Thread.currentThread().isInterrupted()) {'),
    code('    try { Thread.sleep(100); }'),
    code('    catch (InterruptedException e) {'),
    code('        Thread.currentThread().interrupt(); break;'),
    code('    }'),
    code('}'),
  
    h2('5. Пріоритети та демони'),
    code('t.setPriority(Thread.MAX_PRIORITY); // 10'),
    code('t.setDaemon(true); // до start()!'),
  
    h2('6. Синхронізація потоків'),
    para('synchronized-блок або метод: лише один потік одночасно може виконувати захищений блок.'),
    code('synchronized (sharedObject) {'),
    code('    // критична секція'),
    code('}'),
    para('wait() / notify() / notifyAll() — механізм умовного очікування (тільки всередині synchronized):'),
    code('synchronized (lock) {'),
    code('    while (!condition) lock.wait();  // звільняє монітор і чекає'),
    code('    // ... робота ...'),
    code('    lock.notifyAll(); // будить усіх, що чекають на lock'),
    code('}'),
  
    h2('7. Управління через ExecutorService'),
    para('Сучасний підхід: не створювати потоки вручну, а подавати задачі до пулу:'),
    code('ExecutorService es = Executors.newFixedThreadPool(4);'),
    code('Future<?> f = es.submit(() -> doWork());'),
    code('es.shutdown();          // чекати завершення поточних задач'),
    code('es.shutdownNow();       // переривати всі задачі'),
    code('es.awaitTermination(10, TimeUnit.SECONDS);'),
  
    h2('8. Volatile'),
    para('volatile-змінна видима всім потокам без кешування. Гарантує видимість, але не атомарність складних операцій:'),
    code('volatile boolean running = true;'),
    code('// потік 1: running = false;'),
    code('// потік 2: while (running) { ... } — побачить зміну'),
  
    pageBreak(),
  ]);
  
  // ── Q17 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 17. Блокування потоку'),
    divider(),
  
    h2('1. Поняття блокування'),
    para('Блокування (blocking) — стан, коли потік призупиняє виконання у очікуванні певної умови: доступності ресурсу, сигналу від іншого потоку, закінчення таймауту. JVM підтримує кілька видів блокування.'),
  
    h2('2. Монітор і synchronized — BLOCKED'),
    para('Коли потік намагається увійти до synchronized-блоку/методу, що вже зайнятий іншим потоком — він переходить у стан BLOCKED і стоїть у черзі монітора (entry set):'),
    code('synchronized (obj) { /* критична секція */ }'),
    para('Щойно власник монітора виходить зі synchronized, ОС планувальник обирає одного з BLOCKED-потоків і переводить його в RUNNABLE.'),
  
    h2('3. wait() — WAITING'),
    para('Потік, що викликає wait(), звільняє монітор і переходить у WAITING. Він не споживає CPU. Прокидається лише після notify()/notifyAll() від іншого потоку в тому ж моніторі:'),
    code('synchronized (queue) {'),
    code('    while (queue.isEmpty()) queue.wait(); // WAITING'),
    code('    Item item = queue.poll();'),
    code('}'),
    code('// В іншому потоці:'),
    code('synchronized (queue) {'),
    code('    queue.add(item);'),
    code('    queue.notify(); // пробудити одного'),
    code('}'),
  
    h2('4. sleep() — TIMED_WAITING'),
    para('Thread.sleep(ms) призупиняє потік на заданий час. На відміну від wait() — НЕ звільняє монітори. Використовується для затримок, ретраїв, симуляцій:'),
    code('try { Thread.sleep(2000); } catch (InterruptedException e) { ... }'),
  
    h2('5. join() — WAITING / TIMED_WAITING'),
    para('t.join() блокує поточний потік до завершення потоку t. join(ms) — з таймаутом:'),
    code('Thread worker = new Thread(() -> heavyComputation());'),
    code('worker.start();'),
    code('worker.join(); // main чекає завершення worker'),
  
    h2('6. Lock та Condition (java.util.concurrent.locks)'),
    para('ReentrantLock — більш гнучка альтернатива synchronized: підтримує спробу захоплення з таймаутом, переривання під час очікування, справедлива черга (fairness):'),
    code('ReentrantLock lock = new ReentrantLock();'),
    code('Condition notEmpty = lock.newCondition();'),
    code('lock.lock();'),
    code('try {'),
    code('    while (queue.isEmpty()) notEmpty.await(); // аналог wait()'),
    code('    process(queue.poll());'),
    code('} finally { lock.unlock(); }'),
    code(''),
    code('// Спроба захоплення з таймаутом:'),
    code('if (lock.tryLock(500, TimeUnit.MILLISECONDS)) {'),
    code('    try { /* ... */ } finally { lock.unlock(); }'),
    code('} else { /* не вдалося за 500 мс */ }'),
  
    h2('7. Deadlock — взаємне блокування'),
    para('Deadlock: два або більше потоків чекають один одного — жоден не може продовжити. Виникає при захопленні кількох замків у різному порядку:'),
    code('// Потік 1: lock(A), потім lock(B)'),
    code('// Потік 2: lock(B), потім lock(A) → DEADLOCK'),
    para('Запобігання: завжди захоплювати замки в одному порядку; використовувати tryLock() з таймаутом; мінімізувати кількість одночасно утримуваних замків.'),
  
    h2('8. Livelock та Starvation'),
    bullet('Livelock: потоки постійно реагують один на одного, але не рухаються вперед'),
    bullet('Starvation: потік ніколи не отримує CPU через постійне витіснення потоками з вищим пріоритетом'),
  
    pageBreak(),
  ]);
  
  // ── Q19 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 19. Синхронізовані методи'),
    divider(),
  
    h2('1. Поняття синхронізації'),
    para('Синхронізація — механізм упорядкування доступу кількох потоків до спільних ресурсів. Без неї виникають стан гонки (race condition): непередбачуваний результат залежно від порядку виконання операцій.'),
    para('Класичний приклад race condition:'),
    code('// Два потоки паралельно: counter++ (три операції: read, increment, write)'),
    code('int counter = 0;'),
    code('// Очікуємо: 200000, отримуємо: 187432 (race condition!)'),
  
    h2('2. synchronized методи'),
    para('Ключове слово synchronized на методі використовує об\'єкт this (для нестатичних) або Class (для статичних) як монітор. Лише один потік може виконувати синхронізований метод об\'єкта одночасно:'),
    code('public class Counter {'),
    code('    private int count = 0;'),
    code(''),
    code('    public synchronized void increment() { count++; }'),
    code('    public synchronized void decrement() { count--; }'),
    code('    public synchronized int get() { return count; }'),
    code('}'),
  
    h2('3. synchronized блоки'),
    para('Більш гнучкий варіант: захищається тільки необхідна частина коду, можна вказати довільний об\'єкт-монітор:'),
    code('public class BankAccount {'),
    code('    private double balance;'),
    code('    private final Object lock = new Object();'),
    code(''),
    code('    public void transfer(double amount) {'),
    code('        synchronized (lock) {'),
    code('            balance += amount;'),
    code('        }'),
    code('        // несинхронізований код тут — паралельно виконується'),
    code('    }'),
    code('}'),
  
    h2('4. Статичні synchronized методи'),
    para('Статичний synchronized-метод захоплює монітор класу (Counter.class), а не об\'єкта. Блокує всі потоки, що звертаються до цього методу будь-якого екземпляра:'),
    code('public static synchronized int nextId() { return ++idCounter; }'),
  
    h2('5. Атомарні класи як альтернатива'),
    para('Для простих операцій (лічильники) java.util.concurrent.atomic надає lock-free реалізації на основі CAS (Compare-And-Swap):'),
    code('AtomicInteger counter = new AtomicInteger(0);'),
    code('counter.incrementAndGet(); // атомарно, без synchronized'),
    code('counter.compareAndSet(5, 10); // якщо == 5, встановити 10'),
    para('Атомарні класи: AtomicInteger, AtomicLong, AtomicBoolean, AtomicReference, AtomicIntegerArray.'),
  
    h2('6. Модель пам\'яті Java (JMM)'),
    para('Java Memory Model визначає коли зміни однієї нитки видимі іншій. synchronized гарантує: перед виходом зі synchronized — всі записи вивантажуються в головну пам\'ять; при вході — читаються свіжі значення. volatile — видимість без взаємного виключення. Без цих засобів потоки можуть бачити застарілі значення з власного кешу.'),
  
    h2('7. Collections.synchronizedXxx'),
    code('List<String> safeList = Collections.synchronizedList(new ArrayList<>());'),
    code('// АЛЕ: iteration не є атомарною! Потрібен explicit synchronized:'),
    code('synchronized (safeList) {'),
    code('    for (String s : safeList) process(s);'),
    code('}'),
    para('Для concurrent-колекцій краще використовувати ConcurrentHashMap, CopyOnWriteArrayList з java.util.concurrent.'),
  
    pageBreak(),
  ]);
  
  // ── Q21 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 21. Високорівневі способи управління потоками: java.util.concurrent'),
    divider(),
  
    h2('1. Чому java.util.concurrent?'),
    para('До Java 5 єдиним засобом синхронізації були synchronized/wait/notify — низькорівневі і схильні до помилок. Пакет java.util.concurrent (JSR-166, Doug Lea) надає надійні, протестовані абстракції вищого рівня.'),
  
    h2('2. Locks (java.util.concurrent.locks)'),
    bullet('ReentrantLock — заміна synchronized з розширеними можливостями'),
    bullet('ReentrantReadWriteLock — кілька читачів або один записувач'),
    bullet('StampedLock — оптимістичне читання (Java 8+)'),
    code('ReadWriteLock rwLock = new ReentrantReadWriteLock();'),
    code('rwLock.readLock().lock();  // кілька потоків читають одночасно'),
    code('rwLock.writeLock().lock(); // монопольний запис'),
  
    h2('3. Semaphore'),
    para("Обмежує кількість потоків, що одночасно отримують доступ до ресурсу. Корисний для пулів з'єднань, rate-limiting:"),
    code('Semaphore sem = new Semaphore(3); // 3 дозволи'),
    code('sem.acquire(); // зменшити на 1, або чекати'),
    code('try { useResource(); } finally { sem.release(); }'),
  
    h2('4. CountDownLatch'),
    para('Дозволяє потоку чекати, поки N інших завершать свою роботу:'),
    code('CountDownLatch latch = new CountDownLatch(5);'),
    code('for (int i = 0; i < 5; i++) {'),
    code('    executor.submit(() -> { doWork(); latch.countDown(); });'),
    code('}'),
    code('latch.await(); // чекати поки всі 5 завершать'),
  
    h2('5. CyclicBarrier'),
    para('На відміну від CountDownLatch — можна використовувати повторно. Усі потоки зустрічаються в "бар\'єрній точці":'),
    code('CyclicBarrier barrier = new CyclicBarrier(4, () -> System.out.println("All arrived!"));'),
    code('// Кожен потік:'),
    code('doPhase1();'),
    code('barrier.await(); // чекати всіх на бар\'єрі'),
    code('doPhase2();'),
  
    h2('6. Exchanger'),
    para('Два потоки обмінюються об\'єктами в точці зустрічі:'),
    code('Exchanger<Data> exchanger = new Exchanger<>();'),
    code('// Потік 1: Data d2 = exchanger.exchange(myData1);'),
    code('// Потік 2: Data d1 = exchanger.exchange(myData2);'),
  
    h2('7. Concurrent колекції'),
    bullet('ConcurrentHashMap — потокобезпечна HashMap без глобального lock (lock per segment)'),
    bullet('CopyOnWriteArrayList — копіює при кожному записі; ефективна для частих читань'),
    bullet('LinkedBlockingQueue / ArrayBlockingQueue — потокобезпечні черги для Producer-Consumer'),
    bullet('PriorityBlockingQueue — черга з пріоритетом'),
    bullet('ConcurrentLinkedQueue — lock-free черга'),
  
    h2('8. Atomic змінні'),
    bullet('AtomicInteger, AtomicLong — атомарні CAS-операції'),
    bullet('AtomicReference<V> — атомарне оновлення посилань'),
    bullet('LongAdder, LongAccumulator (Java 8) — ефективніше за AtomicLong при великій конкуренції'),
  
    pageBreak(),
  ]);
  
  // ── Q23 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 23. Пули потоків'),
    divider(),
  
    h2('1. Навіщо пул потоків?'),
    para('Створення та знищення потоків — дорога операція: виділення пам\'яті для стеку (~512KB), реєстрація в ОС, context switch. Пул потоків (Thread Pool) — набір попередньо створених потоків, що очікують задачі у черзі. Задача отримує готовий потік миттєво, після завершення — потік повертається до пулу.'),
    para('Переваги: менше накладних витрат на створення, контрольована кількість потоків, захист від надмірного навантаження.'),
  
    h2('2. ExecutorService та Executors'),
    code('// FixedThreadPool — фіксована кількість потоків'),
    code('ExecutorService fixed = Executors.newFixedThreadPool(4);'),
    code(''),
    code('// CachedThreadPool — динамічний розмір (60с TTL для idle потоків)'),
    code('ExecutorService cached = Executors.newCachedThreadPool();'),
    code(''),
    code('// SingleThreadExecutor — один потік, гарантована послідовність'),
    code('ExecutorService single = Executors.newSingleThreadExecutor();'),
    code(''),
    code('// ScheduledThreadPool — виконання по розкладу'),
    code('ScheduledExecutorService sched = Executors.newScheduledThreadPool(2);'),
    code('sched.scheduleAtFixedRate(task, 0, 1, TimeUnit.SECONDS);'),
  
    h2('3. ThreadPoolExecutor — деталі'),
    para('Під капотом Executors.newFixedThreadPool() створює ThreadPoolExecutor з параметрами:'),
    code('new ThreadPoolExecutor('),
    code('    corePoolSize,        // мінімум "живих" потоків'),
    code('    maximumPoolSize,     // максимум потоків'),
    code('    keepAliveTime,       // час простою понад core'),
    code('    TimeUnit.SECONDS,'),
    code('    new LinkedBlockingQueue<>(), // черга задач'),
    code('    new ThreadFactory(),         // фабрика потоків'),
    code('    new CallerRunsPolicy()       // rejection policy'),
    code(');'),
    para('Rejection Policies (при переповненні черги):'),
    bullet('AbortPolicy (default) — кидає RejectedExecutionException'),
    bullet('CallerRunsPolicy — задачу виконує потік-відправник (природний backpressure)'),
    bullet('DiscardPolicy — мовчки відкидає задачу'),
    bullet('DiscardOldestPolicy — відкидає найстарішу задачу з черги'),
  
    h2('4. FixedThreadPool для системи масового обслуговування'),
    para('Класичне застосування: симуляція черги обслуговування (M/M/c). N клієнтів генерує Producer, c потоків-серверів обробляє:'),
    code('ExecutorService servers = Executors.newFixedThreadPool(NUM_SERVERS);'),
    code('BlockingQueue<Customer> queue = new ArrayBlockingQueue<>(MAX_QUEUE);'),
    code(''),
    code('// Producer:'),
    code('while (true) {'),
    code('    Customer c = new Customer(arrivalTime);'),
    code('    queue.put(c); // блокує якщо черга повна'),
    code('    Thread.sleep(poissonDelay(lambda));'),
    code('}'),
    code('// Consumer (кожен сервер):'),
    code('servers.submit(() -> {'),
    code('    while (true) {'),
    code('        Customer c = queue.take(); // блокує якщо черга порожня'),
    code('        Thread.sleep(serviceTime(mu));'),
    code('        recordStats(c);'),
    code('    }'),
    code('});'),
  
    h2('5. Завершення пулу'),
    code('pool.shutdown();                          // чекати завершення поточних задач'),
    code('pool.awaitTermination(30, TimeUnit.SECONDS);'),
    code('// Або примусово:'),
    code('List<Runnable> pending = pool.shutdownNow(); // interrupt всіх'),
  
    pageBreak(),
  ]);
  
  // ── Q25 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 25. ForkJoinFramework: розробка паралельних програм'),
    divider(),
  
    h2('1. Ідея Fork/Join'),
    para('ForkJoinPool (Java 7+) реалізує парадигму "розділяй і володарюй" для паралельних обчислень. Задача рекурсивно розбивається (fork) на підзадачі, які виконуються паралельно, потім результати об\'єднуються (join).'),
    para('Ключова відмінність від ThreadPool: Work Stealing (крадіжка роботи). Якщо потік виконав усі свої задачі, він "краде" задачі з черги сусіднього потоку. Це забезпечує природний балансування навантаження без явного управління.'),
  
    h2('2. RecursiveTask vs RecursiveAction'),
    bullet('RecursiveTask<V> — повертає результат (аналог Callable)'),
    bullet('RecursiveAction — не повертає результат (аналог Runnable)'),
  
    h2('3. RecursiveTask — паралельне підсумовування'),
    code('class ParallelSum extends RecursiveTask<Long> {'),
    code('    static final int THRESHOLD = 10_000;'),
    code('    private final long[] array;'),
    code('    private final int lo, hi;'),
    code(''),
    code('    ParallelSum(long[] a, int lo, int hi) {'),
    code('        this.array = a; this.lo = lo; this.hi = hi;'),
    code('    }'),
    code(''),
    code('    @Override'),
    code('    protected Long compute() {'),
    code('        if (hi - lo <= THRESHOLD) { // базовий випадок'),
    code('            long sum = 0;'),
    code('            for (int i = lo; i < hi; i++) sum += array[i];'),
    code('            return sum;'),
    code('        }'),
    code('        int mid = (lo + hi) / 2;'),
    code('        ParallelSum left = new ParallelSum(array, lo, mid);'),
    code('        ParallelSum right = new ParallelSum(array, mid, hi);'),
    code('        left.fork();                   // асинхронно'),
    code('        long rightResult = right.compute(); // синхронно'),
    code('        return rightResult + left.join(); // чекати left'),
    code('    }'),
    code('}'),
    code('// Виконання:'),
    code('ForkJoinPool pool = ForkJoinPool.commonPool();'),
    code('long sum = pool.invoke(new ParallelSum(array, 0, array.length));'),
  
    h2('4. RecursiveAction — пошук у групі файлів'),
    para('Класичне використання: рекурсивний обхід дерева директорій із пошуком файлів за шаблоном:'),
    code('class FileSearchTask extends RecursiveAction {'),
    code('    private final File dir;'),
    code('    private final String pattern;'),
    code('    private final List<File> results;'),
    code(''),
    code('    @Override'),
    code('    protected void compute() {'),
    code('        File[] entries = dir.listFiles();'),
    code('        if (entries == null) return;'),
    code('        List<FileSearchTask> subTasks = new ArrayList<>();'),
    code('        for (File f : entries) {'),
    code('            if (f.isDirectory()) {'),
    code('                FileSearchTask sub = new FileSearchTask(f, pattern, results);'),
    code('                subTasks.add(sub);'),
    code('                sub.fork(); // паралельно'),
    code('            } else if (f.getName().contains(pattern)) {'),
    code('                results.add(f); // потокобезпечний список!'),
    code('            }'),
    code('        }'),
    code('        for (FileSearchTask t : subTasks) t.join();'),
    code('    }'),
    code('}'),
    code('// results = new CopyOnWriteArrayList<>(); — потокобезпечний список'),
  
    h2('5. Вибір THRESHOLD'),
    para('Threshold — мінімальний розмір підзадачі для рекурсії. Якщо задача менше порогу — вирішується послідовно. Занадто малий поріг: overhead від fork/join перевищує вигоду. Занадто великий: недостатній паралелізм. Правило: порогова задача повинна виконуватись ~0.1–1 мс.'),
  
    h2('6. commonPool'),
    para('ForkJoinPool.commonPool() — глобальний пул JVM з числом потоків = Runtime.getRuntime().availableProcessors() - 1. Спільно використовується parallelStream(), CompletableFuture тощо. Не слід блокувати його потоки довго.'),
  
    pageBreak(),
  ]);
  
  // ── Q27 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 27. Модель обчислень у вигляді графа «операції–операнди»'),
    divider(),
  
    h2('1. Що таке граф операції-операнди'),
    para('Граф операції-операнди (Operation-Operand Graph, або граф залежностей даних) — це спрямований ациклічний граф (DAG), де вузли — операції або операнди (дані), дуги — залежності. Дуга від вузла A до вузла B означає: B не може починатись, поки A не завершиться.'),
  
    h2('2. Призначення моделі'),
    para('Модель відповідає на питання: які операції можна виконувати паралельно? Операції, між якими немає шляху в DAG, не залежать одна від одної і можуть виконуватись одночасно.'),
  
    h2('3. Приклад: обчислення виразу'),
    para('Вираз: z = (a+b) * (c-d) + e'),
    bullet('op1: t1 = a + b'),
    bullet('op2: t2 = c - d'),
    bullet('op3: t3 = t1 * t2  (залежить від op1 і op2)'),
    bullet('op4: z = t3 + e    (залежить від op3)'),
    para('op1 і op2 незалежні → можна виконувати паралельно. op3 чекає op1 і op2. op4 чекає op3. Критичний шлях: op1 → op3 → op4 (або op2 → op3 → op4).'),
  
    h2('4. Критичний шлях'),
    para('Критичний шлях (Critical Path) — найдовший шлях у DAG від початкового до кінцевого вузла. Його довжина визначає нижню межу часу виконання паралельної програми незалежно від кількості процесорів. T_∞ = довжина критичного шляху.'),
  
    h2('5. Паралелізм алгоритму'),
    para('Ступінь паралелізму = T₁ / T_∞, де T₁ — час послідовного виконання, T_∞ — критичний шлях. Наприклад, при T₁=100, T_∞=20 — паралелізм = 5 (максимальне корисне число процесорів = 5).'),
  
    h2('6. Типи залежностей'),
    bullet('RAW (Read After Write) — справжня залежність: op2 читає результат op1'),
    bullet('WAR (Write After Read) — антизалежність: op2 пише туди, звідки op1 читає'),
    bullet('WAW (Write After Write) — залежність виведення: обидва пишуть в одне місце'),
    para('WAR і WAW — "несправжні" залежності, усуваються перейменуванням змінних або функціональним стилем (immutable data).'),
  
    h2('7. Застосування у ForkJoin'),
    para('ForkJoinPool будує неявний DAG задач: fork() створює вузол, join() — дугу залежності. JVM виконує незалежні вузли паралельно через work stealing.'),
  
    pageBreak(),
  ]);
  
  // ── Q29 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 29. Показники ефективності паралельного алгоритму'),
    divider(),
  
    h2('1. Час виконання'),
    bullet('T₁ — час послідовного виконання на 1 процесорі'),
    bullet('Tₚ — час паралельного виконання на p процесорах'),
    bullet('T_∞ — час виконання на ∞ процесорах = критичний шлях'),
  
    h2('2. Прискорення (Speedup)'),
    para('S(p) = T₁ / Tₚ — у скільки разів прискорилось виконання.'),
    bullet('S(p) = p — лінійне прискорення (ідеал, рідко досяжний)'),
    bullet('S(p) < p — реальна ситуація (накладні витрати)'),
    bullet('S(p) > p — суперлінійне (можливо при кращому використанні кешу)'),
  
    h2('3. Ефективність (Efficiency)'),
    para('E(p) = S(p) / p = T₁ / (p × Tₚ). Показує, наскільки добре використовується кожен процесор. E = 1 (100%) — ідеал. E < 1 — є накладні витрати. Типово: E = 0.7–0.9 для добре паралелізованих задач.'),
  
    h2('4. Вартість (Cost)'),
    para('C(p) = p × Tₚ — загальна кількість процесоро-секунд. Оптимально: C(p) = T₁ (не більше, ніж послідовний). Якщо C(p) > T₁ — алгоритм марнує ресурси.'),
  
    h2('5. Масштабованість (Scalability)'),
    para('Слабка масштабованість (weak scaling): при збільшенні p у k разів збільшуємо і розмір задачі у k разів — час має залишатись сталим. Сильна масштабованість (strong scaling): розмір задачі фіксований, збільшуємо p — час зменшується. Ідеальна сильна масштабованість → лінійне прискорення.'),
  
    h2('6. Накладні витрати паралелізму'),
    bullet('Ініціалізація потоків (thread creation overhead)'),
    bullet('Синхронізація (lock contention, barrier synchronization)'),
    bullet('Комунікація (обмін повідомленнями в MPI)'),
    bullet('Незбалансованість навантаження (load imbalance)'),
    bullet('Послідовні частини (закон Амдала)'),
  
    h2('7. Показник ізоефективності'),
    para('Функція ізоефективності W(p) показує, як треба збільшити розмір задачі W при збільшенні p до p\', щоб ефективність залишилась сталою. Чим повільніше зростає W(p), тим більш масштабованим є алгоритм.'),
  
    pageBreak(),
  ]);
  
  // ── Q31 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 31. Закон Амдала. Ефект Амдала'),
    divider(),
  
    h2('1. Формулювання закону Амдала (1967)'),
    para('Закон Амдала (Gene Amdahl) визначає максимальне теоретичне прискорення програми при використанні кількох процесорів. Нехай f — частка програми, яку НЕ можна розпаралелити (послідовна частина), тоді:'),
    para('S(p) = 1 / (f + (1-f)/p)'),
    para('При p → ∞: S_max = 1/f. Тобто, якщо навіть 5% коду є послідовним — максимальне прискорення обмежене 20x, незалежно від кількості процесорів!'),
  
    h2('2. Приклади'),
    bullet('f = 0.50 (50% послідовного) → S_max = 2'),
    bullet('f = 0.25 (25% послідовного) → S_max = 4'),
    bullet('f = 0.10 (10% послідовного) → S_max = 10'),
    bullet('f = 0.05 (5% послідовного)  → S_max = 20'),
    bullet('f = 0.01 (1% послідовного)  → S_max = 100'),
  
    h2('3. Ефект Амдала'),
    para('При збільшенні p від 1 до 8 прискорення зростає значно. Але з 8 до 16, 16 до 32 — приріст все менший. Після певного p додавання процесорів майже нічого не дає. Це "ефект Амдала" або "стеля прискорення".'),
    para('Висновок: для отримання значного прискорення необхідно мінімізувати послідовну частину. Навіть невеликий відсоток послідовного коду критично обмежує масштабованість.'),
  
    h2('4. Java-приклад вимірювання прискорення'),
    code('long start = System.nanoTime();'),
    code('// послідовно:'),
    code('long seqResult = sequentialSum(data);'),
    code('long T1 = System.nanoTime() - start;'),
    code(''),
    code('start = System.nanoTime();'),
    code('// паралельно на p ядрах:'),
    code('long parResult = parallelSum(data, p);'),
    code('long Tp = System.nanoTime() - start;'),
    code(''),
    code('double speedup = (double) T1 / Tp;'),
    code('double efficiency = speedup / p;'),
    code('System.out.printf("S(%d) = %.2f, E(%d) = %.2f%%%n", p, speedup, p, efficiency*100);'),
  
    h2('5. Закон Густафсона (розширення Амдала)'),
    para('Густафсон (1988) переформулював: при збільшенні p можна збільшувати розмір задачі (weak scaling). Тоді: S(p) = p - f*(p-1). При великих p прискорення може бути значно більшим, ніж передбачає Амдал, якщо дозволяти збільшувати задачу. Це обґрунтування практичної корисності суперкомп\'ютерів.'),
  
    h2('6. Практичні наслідки'),
    bullet('Профілюйте перед паралелізацією: знайдіть справжню послідовну частину (f)'),
    bullet('Зменшуйте f: рефакторуйте ввід/вивід, ініціалізацію, синхронізацію'),
    bullet('Не марнуйте ресурси: якщо f=20%, купівля 100 ядер — марна трата грошей'),
    bullet('Weak scaling може бути кращою стратегією, ніж strong scaling'),
  
    pageBreak(),
  ]);
  
  // ── Q33 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 33. Оцінка ефективності паралельних обчислень'),
    divider(),
  
    h2('1. Методологія вимірювання'),
    para('Для об\'єктивної оцінки потрібно: виміряти T₁ (послідовний алгоритм) і Tₚ (паралельний на p ядрах); варіювати p і розмір задачі N; усереднити по кількох запусках; виключити JVM-прогрів (warmup).'),
    code('// Правильний benchmarking у Java:'),
    code('// 1) JVM warmup — кілька "холодних" запусків'),
    code('for (int w = 0; w < 5; w++) parallelSum(data, p);'),
    code('// 2) Реальне вимірювання'),
    code('long[] times = new long[10];'),
    code('for (int r = 0; r < 10; r++) {'),
    code('    long t = System.nanoTime();'),
    code('    parallelSum(data, p);'),
    code('    times[r] = System.nanoTime() - t;'),
    code('}'),
    code('// 3) Медіана або середнє'),
  
    h2('2. Аналіз прискорення (Speedup curve)'),
    para('Будується графік S(p) для різних значень p. Ідеальна пряма S(p)=p. Реальна крива відхиляється вгору (при малих p) і виходить на плато (ефект Амдала). Де крива відхиляється від лінійної — є накладні витрати.'),
  
    h2('3. Аналіз ефективності'),
    para('E(p) = S(p)/p. Для добрих алгоритмів E(p) ≥ 0.7 при помірних p. Різке падіння E при збільшенні p вказує на: надмірну синхронізацію, дисбаланс навантаження, комунікаційне вузьке місце.'),
  
    h2('4. Профілювання'),
    para('Інструменти для Java: JProfiler, YourKit, VisualVM, Java Flight Recorder (JFR). Дозволяють знайти: де потоки заблоковані (lock contention), які методи займають найбільше часу, стан потоків у часі.'),
  
    h2('5. Оцінка для розподілених систем (MPI)'),
    para('Для MPI-програм додатково аналізують: T_comm — час комунікацій, T_comp — час обчислень, співвідношення T_comp / T_comm (має бути >> 1 для ефективності), балансування навантаження між процесами (через MPI_Gather статистики).'),
  
    h2('6. Мікробенчмарки — JMH'),
    para('Java Microbenchmark Harness (JMH) — стандартний інструмент для точного вимірювання продуктивності Java-коду з правильним врахуванням JIT-компіляції та прогріву:'),
    code('@Benchmark'),
    code('@BenchmarkMode(Mode.AverageTime)'),
    code('@OutputTimeUnit(TimeUnit.MILLISECONDS)'),
    code('public long benchmarkParallelSum(Blackhole bh) {'),
    code('    return parallelSum(data, 4);'),
    code('}'),
  
    pageBreak(),
  ]);
  
  // ── Q35 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 35. Архітектура розподілених систем. Таксономія Фліна'),
    divider(),
  
    h2('1. Таксономія Фліна (Flynn, 1966)'),
    para('Класифікація комп\'ютерних архітектур за потоками команд (Instruction Streams) та потоками даних (Data Streams):'),
  
    h2('SISD — Single Instruction, Single Data'),
    para('Класичний послідовний процесор. Одна команда обробляє один операнд. Приклад: однопроцесорний комп\'ютер без векторних розширень.'),
  
    h2('SIMD — Single Instruction, Multiple Data'),
    para('Одна команда застосовується до вектора даних одночасно. Приклади: векторні розширення x86 (SSE, AVX, AVX-512), GPU-ядра, DSP-процесори. Використовується для мультимедіа, ML, наукових обчислень.'),
  
    h2('MISD — Multiple Instruction, Single Data'),
    para('Кілька команд обробляють одні дані. Рідко використовується на практиці. Прикладом є конвеєрна обробка або резервовані системи (fault-tolerant).'),
  
    h2('MIMD — Multiple Instruction, Multiple Data'),
    para('Кілька процесорів виконують різні команди над різними даними незалежно. Найпоширеніша архітектура паралельних та розподілених систем:'),
    bullet('Shared Memory MIMD: багатоядерні CPU, SMP (Symmetric Multiprocessing)'),
    bullet("Distributed Memory MIMD: кластери, суперкомп\'ютери (вузли з'єднані через MPI)"),
  
    h2('2. Архітектура розподілених систем'),
    para('Розподілена система — сукупність комп\'ютерів, що виглядають для користувача як єдина система. Ключові характеристики: паралельність, відсутність спільного годинника, незалежні відмови вузлів.'),
  
    h2('3. Класифікація за пам\'яттю'),
    bullet('UMA (Uniform Memory Access) — SMP, всі процесори рівно далеко від RAM'),
    bullet('NUMA (Non-Uniform Memory Access) — локальна RAM швидша за чужу (сучасні многосокетні сервери)'),
    bullet('COMA (Cache-Only Memory Architecture) — уся пам\'ять є кешем'),
    bullet('MPP (Massively Parallel Processing) — кожен вузол має власну RAM, зв\'язок через мережу'),
  
    h2('4. Мережі міжз\'єднань'),
    bullet('Bus — загальна шина (обмежена пропускна здатність)'),
    bullet('Ring — кільцева топологія'),
    bullet('Mesh / Torus — 2D/3D решітка (Blue Gene, Cray)'),
    bullet('Hypercube — N-мірний куб (2ᴺ вузлів, діаметр N)'),
    bullet('Fat-Tree / InfiniBand — сучасні HPC-кластери'),
  
    pageBreak(),
  ]);
  
  // ── Q37 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 37. Стандарти та технології обміну повідомленнями в розподілених системах'),
    divider(),
  
    h2('1. Message Passing Interface (MPI)'),
    para('MPI — стандарт інтерфейсу передачі повідомлень для розробки паралельних програм на кластерах. Версія 1.0 — 1994, MPI-3.1 — актуальна. Реалізації: OpenMPI, MPICH, Intel MPI. Підтримує C, C++, Fortran; через обгортки — Python (mpi4py), Java (MPJ Express).'),
    para('Модель: SPMD (Single Program Multiple Data) — всі процеси запускають одну програму, але виконують різні гілки залежно від rank. Кожен процес має свою пам\'ять; взаємодія — лише через явні виклики MPI.'),
  
    h2('2. Основні поняття MPI'),
    bullet('Communicator (MPI_COMM_WORLD) — група процесів'),
    bullet('Rank — унікальний номер процесу в комунікаторі (0..size-1)'),
    bullet('Tag — числовий тег повідомлення для розрізнення'),
    bullet('Point-to-Point: MPI_Send / MPI_Recv (між двома процесами)'),
    bullet('Collective: MPI_Bcast, MPI_Scatter, MPI_Gather, MPI_Reduce (всі процеси)'),
  
    h2('3. Apache Kafka — обмін повідомленнями для великих систем'),
    para('Kafka — розподілена платформа потокової передачі даних. Продюсери публікують у топіки, консюмери читають з партицій. Масштабується до мільйонів повідомлень/сек. Використовується в мікросервісах замість прямих HTTP-викликів.'),
  
    h2('4. gRPC / Protocol Buffers'),
    para('Google\'s Remote Procedure Call: бінарний протокол на основі HTTP/2 та Protobuf. Швидший за REST/JSON у 5-10 разів. Підтримує streaming. Широко використовується в мікросервісних архітектурах (Spring Boot + gRPC).'),
  
    h2('5. REST / HTTP'),
    para('Найпростіший і найпоширеніший підхід для web-сервісів. JSON або XML payload. Синхронний request-response. Простота і підтримка браузерів. Менш ефективний за Kafka/gRPC для high-throughput систем.'),
  
    h2('6. RabbitMQ / ActiveMQ — Message Brokers'),
    para('Message broker: посередник між продюсером і консюмером. Підтримує черги, маршрутизацію, публікацію/підписку. AMQP-протокол. Гарантія доставки (acknowledgment). Використовується для асинхронної обробки задач.'),
  
    h2('7. Порівняння для Spring Boot застосунку'),
    bullet('REST: прості CRUD-операції, admin API, прозорість'),
    bullet('gRPC: внутрішні мікросервіси, де важлива швидкість'),
    bullet('Kafka: event-driven архітектура, великі обсяги даних'),
    bullet('MPI: наукові HPC-обчислення (не типово для Spring Boot)'),
  
    pageBreak(),
  ]);
  
  // ── Q39 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 39. Методи обміну повідомленнями MPI один-до-одного та їх застосування'),
    divider(),
  
    h2('1. Блокуючі Point-to-Point операції'),
    para('MPI_Send(buf, count, datatype, dest, tag, comm) — надіслати повідомлення процесу dest. Блокує, поки повідомлення не буде безпечно буферизоване або отримане.'),
    para('MPI_Recv(buf, count, datatype, source, tag, comm, status) — отримати повідомлення від source. Блокує до отримання.'),
    code('// Процес 0 відправляє, процес 1 отримує:'),
    code('if (rank == 0) {'),
    code('    int data = 42;'),
    code('    MPI_Send(&data, 1, MPI_INT, 1, 0, MPI_COMM_WORLD);'),
    code('} else if (rank == 1) {'),
    code('    int data;'),
    code('    MPI_Recv(&data, 1, MPI_INT, 0, 0, MPI_COMM_WORLD, MPI_STATUS_IGNORE);'),
    code('}'),
  
    h2('2. Неблокуючі операції'),
    para('MPI_Isend / MPI_Irecv — повертаються негайно, не чекаючи завершення. Потрібно завершити через MPI_Wait або перевірити через MPI_Test. Дозволяють перекривати комунікацію та обчислення:'),
    code('MPI_Request req;'),
    code('MPI_Isend(sendBuf, N, MPI_DOUBLE, dest, 0, comm, &req);'),
    code('// Поки йде відправка — обчислюємо:'),
    code('localComputation();'),
    code('MPI_Wait(&req, MPI_STATUS_IGNORE); // чекати завершення'),
  
    h2('3. MPI_Sendrecv'),
    para('Атомарна операція: відправити і отримати одночасно. Запобігає дедлокам при кільцевих обмінах (наприклад, зсуви в алгоритмі Кеннона):'),
    code('MPI_Sendrecv(sendBuf, N, MPI_DOUBLE, rightNeighbor, 0,'),
    code('             recvBuf, N, MPI_DOUBLE, leftNeighbor,  0,'),
    code('             comm, MPI_STATUS_IGNORE);'),
  
    h2('4. Режими відправки'),
    bullet('MPI_Send (standard): автоматично вибирає буферизований або синхронний'),
    bullet('MPI_Bsend (buffered): завжди буферизується, повертається відразу'),
    bullet('MPI_Ssend (synchronous): чекає, поки recv не буде викликано'),
    bullet('MPI_Rsend (ready): лише якщо recv вже готовий'),
  
    h2('5. Застосування для множення матриць (стрічковий алгоритм MPI)'),
    code('// Процес 0 розсилає рядки матриці A:'),
    code('MPI_Scatter(A, rowsPerProc * M, MPI_DOUBLE,'),
    code('            localA, rowsPerProc * M, MPI_DOUBLE, 0, comm);'),
    code('// Всі отримують повну матрицю B:'),
    code('MPI_Bcast(B, M*K, MPI_DOUBLE, 0, comm);'),
    code('// Локальне множення'),
    code('multiplyLocal(localA, B, localC, rowsPerProc, M, K);'),
    code('// Збір результатів:'),
    code('MPI_Gather(localC, rowsPerProc * K, MPI_DOUBLE,'),
    code('           C, rowsPerProc * K, MPI_DOUBLE, 0, comm);'),
  
    pageBreak(),
  ]);
  
  // ── Q41 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 41. OpenMPI: особливості побудови програми та основні методи'),
    divider(),
  
    h2('1. Що таке OpenMPI'),
    para('OpenMPI — відкрита реалізація стандарту MPI. Найпопулярніша у HPC-середовищах. Підтримує: TCP, InfiniBand, Shared Memory транспорти; динамічні процеси (MPI_Spawn); fault tolerance.'),
  
    h2('2. Структура MPI-програми'),
    code('#include <mpi.h>'),
    code('int main(int argc, char** argv) {'),
    code('    MPI_Init(&argc, &argv);           // ініціалізація'),
    code('    int rank, size;'),
    code('    MPI_Comm_rank(MPI_COMM_WORLD, &rank); // номер процесу'),
    code('    MPI_Comm_size(MPI_COMM_WORLD, &size); // кількість процесів'),
    code(''),
    code('    if (rank == 0) { /* master */ }'),
    code('    else { /* worker */ }'),
    code(''),
    code('    MPI_Finalize();                   // завершення'),
    code('    return 0;'),
    code('}'),
    para('Компіляція та запуск:'),
    code('mpicc -o program program.c'),
    code('mpirun -np 4 ./program          # 4 процеси на 1 вузлі'),
    code('mpirun -np 16 --hostfile hosts ./program  # 16 процесів на кластері'),
  
    h2('3. Колективні методи OpenMPI'),
    bullet('MPI_Bcast(buf, count, type, root, comm) — розсилає від root до всіх'),
    bullet('MPI_Scatter(sendbuf, cnt, type, recvbuf, cnt, type, root, comm) — розподіляє рівні частини від root'),
    bullet('MPI_Gather(sendbuf, cnt, type, recvbuf, cnt, type, root, comm) — збирає від усіх до root'),
    bullet('MPI_Allgather — Gather, але результат у всіх'),
    bullet('MPI_Reduce(sendbuf, recvbuf, cnt, type, op, root, comm) — редукція (сума, макс тощо) до root'),
    bullet('MPI_Allreduce — Reduce, але результат у всіх'),
    bullet('MPI_Barrier(comm) — синхронізаційний бар\'єр'),
    bullet('MPI_Alltoall — транспозиція: кожен надсилає кожному рівну частину'),
  
    h2('4. Операції редукції'),
    para('Вбудовані: MPI_SUM, MPI_PROD, MPI_MAX, MPI_MIN, MPI_LAND (логічне AND), MPI_LOR, MPI_BAND (побітове AND). Можна визначати власні через MPI_Op_create.'),
  
    h2('5. Комунікатори'),
    para('Комунікатор визначає групу процесів. MPI_COMM_WORLD — всі процеси. Можна створювати підгрупи для структурованих алгоритмів:'),
    code('MPI_Comm rowComm;'),
    code('MPI_Comm_split(MPI_COMM_WORLD, myRow, myCol, &rowComm);'),
    code('// Тепер MPI_Bcast в rowComm — тільки по рядку решітки'),
  
    h2('6. Типи даних MPI'),
    bullet('MPI_INT, MPI_DOUBLE, MPI_FLOAT, MPI_CHAR, MPI_BYTE'),
    bullet('MPI_Type_contiguous, MPI_Type_vector — похідні типи для нерегулярних даних'),
    bullet('MPI_Type_struct — структури з полями різних типів'),
  
    pageBreak(),
  ]);
  
  // ── Q43 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 43. Алгоритми паралельного сортування'),
    divider(),
  
    h2('1. Навіщо паралельне сортування?'),
    para('Послідовне сортування O(N log N) стає вузьким місцем при великих N (мільярди елементів). Паралельне сортування дозволяє знизити час до O(N log N / p).'),
  
    h2('2. Паралельне злиттям (Parallel Merge Sort)'),
    para('Розбити масив на p рівних частин → кожен потік/процес сортує свою частину → злити відсортовані частини. Злиття p частин займає O(p * N/p) = O(N) послідовно або O(N/p * log p) паралельно.'),
    code('class MergeSortTask extends RecursiveAction {'),
    code('    protected void compute() {'),
    code('        if (size <= THRESHOLD) { Arrays.sort(data, lo, hi); return; }'),
    code('        MergeSortTask left = new MergeSortTask(data, lo, mid);'),
    code('        MergeSortTask right = new MergeSortTask(data, mid, hi);'),
    code('        invokeAll(left, right);'),
    code('        merge(data, lo, mid, hi);'),
    code('    }'),
    code('}'),
  
    h2('3. Паралельний QuickSort'),
    para('Вибрати pivot → partition → рекурсивно сортувати ліву і праву частини паралельно. Проблема: якщо pivot поганий — нерівний розподіл. Вирішення: медіана трьох або рандомний pivot.'),
  
    h2('4. Bitonic Sort'),
    para('Сортувальна мережа: послідовність компараторів (порівняти-замінити), яка не залежить від даних. Паралельно: O(log²N) кроків по O(N/2) паралельних порівнянь. Ефективний для GPU і FPGA.'),
  
    h2('5. Паралельне сортування в Java'),
    code('// Arrays.parallelSort() — вбудовано з Java 8, використовує ForkJoin:'),
    code('int[] array = new int[1_000_000];'),
    code('Arrays.parallelSort(array);'),
    code(''),
    code('// Паралельне сортування частини:'),
    code('Arrays.parallelSort(array, 100, 500);'),
  
    h2('6. Сортування злиттям у MPI'),
    para('N елементів на p процесах: Scatter рівних частин → локальне sort → попарне злиття: на кожному кроці половина процесів відправляє свою частину сусіду, сусід зливає і продовжує. log₂(p) кроків.'),
    bullet('Крок 0: 0 merge з 1, 2 merge з 3, 4 merge з 5...'),
    bullet('Крок 1: 0 merge з 2, 4 merge з 6...'),
    bullet('Крок log₂p: процес 0 має повністю відсортований масив'),
  
    h2('7. Паралельний Java 8 parallelStream sort'),
    code('List<Integer> sorted = list.parallelStream()'),
    code('    .sorted()'),
    code('    .collect(Collectors.toList());'),
    para('Під капотом: ForkJoinPool.commonPool() + Arrays.parallelSort.'),
  
    pageBreak(),
  ]);
  
  // ── Q45 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 45. Оцінювання продуктивності паралельних алгоритмів у розподілених системах'),
    divider(),
  
    h2('1. Специфіка розподілених систем'),
    para('На відміну від shared-memory, у distributed-memory системах (MPI-кластери) час виконання = T_comp + T_comm. T_comm може домінувати при великій кількості вузлів або дрібних задачах.'),
  
    h2('2. Модель комунікаційних витрат'),
    para('Час відправки повідомлення розміром m байт: T = α + m/β, де α — латентність (startup time), β — пропускна здатність. InfiniBand: α ≈ 1–2 мкс, β ≈ 100 Гбіт/с. Ethernet: α ≈ 50–100 мкс, β ≈ 10 Гбіт/с.'),
  
    h2('3. Аналіз для стрічкового MPI-алгоритму множення матриць'),
    bullet('MPI_Scatter матриці A: T_scatter = α*log₂p + N²/(2*β)'),
    bullet('MPI_Bcast матриці B: T_bcast = α*log₂p + N²/β'),
    bullet('Локальне множення: T_comp = 2*N³/p'),
    bullet('MPI_Gather результату: T_gather ≈ T_scatter'),
    para('Загально: T_total = 2*N³/p + O(N²*log(p)/β). При великому N переважають обчислення; при малому — комунікації.'),
  
    h2('4. Метрики для кластеру'),
    bullet('Sustained GFlops — реальна продуктивність в Гфлопс'),
    bullet('Peak / Sustained ratio — ефективність використання обладнання'),
    bullet('Parallel efficiency E(p) = T₁/(p*Tₚ)'),
    bullet('Communication overhead fraction = T_comm / T_total'),
  
    h2('5. Інструменти вимірювання'),
    bullet('MPI_Wtime() — таймер з роздільністю ~наносекунда'),
    bullet('Score-P / TAU — профілювання MPI-програм'),
    bullet('Intel VTune, NVIDIA Nsight — аналіз HPC-додатків'),
    bullet('Vampir — візуалізація трасування MPI-комунікацій'),
    code('double t_start = MPI_Wtime();'),
    code('// ... MPI operations ...'),
    code('double t_end = MPI_Wtime();'),
    code('double local_time = t_end - t_start;'),
    code('double max_time;'),
    code('MPI_Reduce(&local_time, &max_time, 1, MPI_DOUBLE, MPI_MAX, 0, comm);'),
  
    h2('6. Балансування навантаження'),
    para('Якщо процеси мають різну тривалість — весь алгоритм чекає найповільнішого (load imbalance). Вирішення: динамічний розподіл задач (work stealing, task queue); нерівномірний поділ при нерівномірних даних (адаптивна декомпозиція).'),
  
    pageBreak(),
  ]);
  
  // ── Q47 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 47. Грід-технології. Поняття спільного віртуального ресурсу'),
    divider(),
  
    h2('1. Що таке Grid Computing'),
    para('Грід-обчислення (Grid Computing) — парадигма розподілених обчислень, де ресурси різних організацій (університетів, дослідницьких інститутів, підприємств) об\'єднуються через Інтернет для вирішення масштабних наукових задач.'),
    para('Термін запропоновано Іаном Фостером і Карлом Кессельманом у 1998 ("The Grid: Blueprint for a New Computing Infrastructure"). Аналогія: електрична мережа (power grid) — ресурс "підключений і використовується, не думаючи про його місцезнаходження".'),
  
    h2('2. Спільний віртуальний ресурс (Virtual Organization)'),
    para('Virtual Organization (VO) — ключова концепція грід: динамічна спільнота людей та ресурсів, що поділяють спільні цілі та правила. Ресурси (CPU, пам\'ять, сховище, програмне забезпечення) різних організацій виглядають як єдиний спільний ресурс, незважаючи на гетерогенність і географічну розподіленість.'),
    para('Приклад: CERN (Женева) + університети 40+ країн → WLCG (Worldwide LHC Computing Grid). Обробляє ~15 петабайт даних щороку від Великого адронного колайдера.'),
  
    h2('3. Рівні архітектури грід (Foster, Kesselman)'),
    bullet('Fabric level — фізичні ресурси (CPU, пам\'ять, файлові системи, мережа)'),
    bullet('Connectivity level — протоколи комунікації та аутентифікації (TCP/IP, PKI, SSL)'),
    bullet('Resource level — протоколи для окремих ресурсів (GRAM — управління задачами, GridFTP — передача файлів)'),
    bullet('Collective level — сервіси для координації ресурсів (брокери ресурсів, планувальники)'),
    bullet('Application level — кінцеві застосунки, що використовують грід'),
  
    h2('4. Характеристики грід-систем'),
    bullet('Гетерогенність: різні ОС, апаратне забезпечення, ПЗ'),
    bullet('Географічна розподіленість: вузли на різних континентах'),
    bullet('Динамічність: вузли можуть приєднуватись і від\'єднуватись'),
    bullet('Безпека: аутентифікація через PKI-сертифікати (X.509)'),
    bullet('Масштабованість: від сотень до мільйонів вузлів'),
  
    h2('5. Відмінності: Грід vs Хмара vs Кластер'),
    bullet('Кластер: однорідні вузли в одному ЦОД, тісна інтеграція'),
    bullet('Грід: гетерогенні вузли різних організацій, слабка інтеграція'),
    bullet('Хмара (Cloud): on-demand ресурси, оплата за використання, повна віртуалізація'),
  
    pageBreak(),
  ]);
  
  // ── Q49 ─────────────────────────────────────────────────────────────────────
  sections.push(...[
    h1('Питання 49. Програмне забезпечення грід-технологій'),
    divider(),
  
    h2('1. Globus Toolkit'),
    para('Де-факто стандарт грід-middleware. Надає: GSI (Grid Security Infrastructure) — аутентифікація через X.509-сертифікати, делегування, єдиний вхід (Single Sign-On); GRAM (Grid Resource Allocation and Management) — подача та управління задачами на вузлах; GridFTP — розширений FTP для надійної передачі великих файлів; MDS (Monitoring and Discovery Service) — реєстр ресурсів; OGSA-DAI — доступ до баз даних.'),
  
    h2('2. BOINC (Berkeley Open Infrastructure for Network Computing)'),
    para('Платформа для волонтерських обчислень: добровольці надають CPU/GPU своїх домашніх комп\'ютерів. Проекти: SETI@Home (пошук позаземного інтелекту), Folding@Home (моделювання білків), Rosetta@Home, Einstein@Home. Мільйони добровольців → петафлопсна потужність без витрат на обладнання.'),
  
    h2('3. HTCondor (раніше Condor)'),
    para('Система управління навантаженням (workload management) для великих колекцій розподілених комп\'ютерів. Особливість: використовує "простої" комп\'ютерів (idle cycles) — якщо комп\'ютер не використовується, HTCondor запускає на ньому задачу. Використовується в WLCG, університетських кластерах.'),
  
    h2('4. gLite / DIRAC'),
    para('gLite — middleware для EGEE (Enabling Grids for E-sciencE) — найбільший науковий грід Європи. DIRAC — новіша розподілена система управління задачами для LHCb (CERN). Надає: workload management, data management, monitoring.'),
  
    h2('5. Apache Hadoop та Spark — "нове покоління грід"'),
    para('Для задач big data: Hadoop HDFS — розподілена файлова система (replicated blocks), Hadoop MapReduce — паралельна обробка (map + reduce + shuffle), Apache Spark — in-memory обчислення, набагато швидший за MapReduce для ітеративних алгоритмів (ML, графи).'),
    code('// Spark Word Count (Java):'),
    code('JavaRDD<String> lines = spark.read().textFile("hdfs://...").javaRDD();'),
    code('JavaPairRDD<String, Integer> counts = lines'),
    code('    .flatMap(l -> Arrays.asList(l.split(" ")).iterator())'),
    code('    .mapToPair(w -> new Tuple2<>(w, 1))'),
    code('    .reduceByKey((a, b) -> a + b);'),
    code('counts.saveAsTextFile("hdfs://output");'),
  
    h2('6. Планувальники задач для кластерів'),
    bullet('SLURM (Simple Linux Utility for Resource Management) — найпопулярніший у HPC'),
    bullet('PBS/Torque — старіший стандарт'),
    bullet('LSF (IBM Spectrum LSF) — комерційний'),
    bullet('Kubernetes — для контейнеризованих задач (cloud-native grid)'),
    code('# SLURM приклад запуску MPI-задачі:'),
    code('#!/bin/bash'),
    code('#SBATCH --nodes=4'),
    code('#SBATCH --ntasks-per-node=16'),
    code('#SBATCH --time=02:00:00'),
    code('mpirun -np 64 ./matrix_multiply 2048'),
  
    h2('7. Безпека в грід'),
    para('GSI (Grid Security Infrastructure): X.509-сертифікати для ідентифікації; proxy-сертифікати для делегування повноважень; VOMS (Virtual Organization Membership Service) — управління членством у VO. Дозволяє безпечно запускати задачі на ресурсах, яким не повністю довіряєш.'),
  
    pageBreak(),
  ]);
  
  // ─── assemble document ───────────────────────────────────────────────────────
  
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Arial', size: 24 } },
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial', color: '1F3864' },
          paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // ~2cm
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 1000, after: 400 },
          children: [new TextRun({ text: 'Технологія паралельних обчислень', bold: true, size: 48, font: 'Arial', color: '1F3864' })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 200 },
          children: [new TextRun({ text: 'Відповіді до екзамену', size: 32, font: 'Arial', color: '2E75B6' })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 1200 },
          children: [new TextRun({ text: 'Java · OpenMPI · ForkJoin · Spring Boot', size: 24, font: 'Arial', italics: true, color: '595959' })],
        }),
        pageBreak(),
        ...sections,
      ],
    }],
  });
  
  Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('./docs/uv5.docx', buffer);
  }).catch(err => {
    console.error('❌ Помилка:', err);
  });