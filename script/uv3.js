// Запустити: node generate_exam_answers.js
// Залежності: npm install docx
// Результат: exam_answers.docx

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageNumber, PageBreak,
    BorderStyle, TabStopType, TabStopPosition
  } = require('docx');
  const fs = require('fs');
  
  // ─────────────────────────────────────────────
  // ДОПОМІЖНІ ФУНКЦІЇ
  // ─────────────────────────────────────────────
  
  function heading1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text, bold: true, size: 32 })],
      spacing: { before: 480, after: 200 },
      pageBreakBefore: true,
    });
  }
  
  function heading2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      children: [new TextRun({ text, bold: true, size: 28 })],
      spacing: { before: 300, after: 120 },
    });
  }
  
  function para(text, options = {}) {
    return new Paragraph({
      children: [new TextRun({ text, size: 24, ...options })],
      spacing: { before: 80, after: 80 },
      alignment: AlignmentType.JUSTIFIED,
    });
  }
  
  function bold(text) {
    return new Paragraph({
      children: [new TextRun({ text, bold: true, size: 24 })],
      spacing: { before: 80, after: 60 },
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      children: [new TextRun({ text, size: 24 })],
      spacing: { before: 40, after: 40 },
    });
  }
  
  function code(text) {
    return new Paragraph({
      children: [new TextRun({ text, font: 'Courier New', size: 20 })],
      spacing: { before: 40, after: 40 },
      indent: { left: 720 },
    });
  }
  
  function separator() {
    return new Paragraph({
      children: [new TextRun({ text: '' })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E75B6' } },
      spacing: { before: 200, after: 200 },
    });
  }
  
  // ─────────────────────────────────────────────
  // ЗМІСТ ВІДПОВІДЕЙ
  // ─────────────────────────────────────────────
  
  const content = [];
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 1
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 1. Поняття паралельних, псевдопаралельних та розподілених обчислень'));
  
  content.push(heading2('1.1. Паралельні обчислення'));
  content.push(para(
    'Паралельні обчислення (Parallel Computing) — це спосіб виконання обчислень, при якому велика задача ' +
    'розбивається на менші підзадачі, що виконуються одночасно на кількох фізичних процесорах або ядрах. ' +
    'Ключова відмінність від псевдопаралелізму — одночасне фізичне виконання інструкцій у один і той же ' +
    'момент часу (true simultaneity). Паралельні обчислення дозволяють суттєво скоротити час розв\'язання ' +
    'задач, що потребують великих обчислювальних ресурсів: чисельне моделювання, обробка зображень, ' +
    'машинне навчання, криптографія тощо.'
  ));
  content.push(para(
    'Технічно паралельні обчислення реалізуються на рівні: процесорних ядер (multi-core CPU), ' +
    'векторних блоків (SIMD — Single Instruction Multiple Data), графічних прискорювачів (GPU), ' +
    'або цілих кластерів з\'єднаних машин.'
  ));
  
  content.push(heading2('1.2. Псевдопаралельні обчислення'));
  content.push(para(
    'Псевдопаралельні обчислення (Pseudo-parallel / Concurrent Computing) — це ілюзія одночасності, ' +
    'що створюється операційною системою на одному процесорі через швидке перемикання між кількома ' +
    'потоками або процесами. Кожен потік отримує невеликий відрізок часу (time slice, квант часу), ' +
    'після якого відбувається перемикання контексту (context switch) до іншого потоку.'
  ));
  content.push(para(
    'Користувачу здається, що програми виконуються одночасно, проте насправді виконується лише одна ' +
    'інструкція за раз. Псевдопаралелізм — це переважно механізм підвищення відгуку системи (responsiveness) ' +
    'та спрощення архітектури програм (кожне завдання пишеться як незалежна послідовна програма). ' +
    'У Java псевдопаралелізм виникає, коли кілька потоків виконуються на одному ядрі процесора.'
  ));
  
  content.push(heading2('1.3. Розподілені обчислення'));
  content.push(para(
    'Розподілені обчислення (Distributed Computing) — це парадигма, при якій задача виконується на ' +
    'кількох окремих комп\'ютерах (вузлах), з\'єднаних мережею. Кожен вузол має власну пам\'ять, ' +
    'свій процесор і власне програмне забезпечення. Взаємодія між вузлами відбувається через ' +
    'передавання повідомлень (message passing), а не через спільну пам\'ять.'
  ));
  content.push(para(
    'Розподілені системи мають ряд принципових особливостей: відсутність спільного глобального годинника, ' +
    'ненадійність мережі, можливість часткових відмов (один вузол падає, інші продовжують роботу). ' +
    'Для координації вузлів використовуються такі стандарти та технології, як MPI (Message Passing Interface), ' +
    'RPC, REST/gRPC, Apache Kafka, Apache Spark.'
  ));
  
  content.push(heading2('1.4. Порівняльна таблиця'));
  content.push(bullet('Паралельні: фізичне одночасне виконання, спільна або розподілена пам\'ять, багатоядерний CPU/GPU/кластер'));
  content.push(bullet('Псевдопаралельні: логічна одночасність на одному ядрі, спільна пам\'ять, контекстне перемикання'));
  content.push(bullet('Розподілені: фізично різні машини, передавання повідомлень через мережу, власна пам\'ять кожного вузла'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 3
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 3. Способи розпаралелювання'));
  
  content.push(heading2('3.1. Декомпозиція за даними (Data Decomposition)'));
  content.push(para(
    'Найпоширеніший спосіб: вхідні дані розбиваються на незалежні частини, і кожна частина ' +
    'обробляється окремим потоком або процесом. Типовий приклад — розбиття рядків матриці між потоками ' +
    'при множенні матриць. У Java це реалізується через розбиття масиву на діапазони: ' +
    'потік i обробляє рядки від i*(n/p) до (i+1)*(n/p), де n — кількість рядків, p — кількість потоків.'
  ));
  
  content.push(heading2('3.2. Декомпозиція за функціями (Functional / Task Decomposition)'));
  content.push(para(
    'Задача розбивається на незалежні підзадачі (функції), кожна з яких виконується окремим потоком. ' +
    'Наприклад, один потік читає дані, другий обробляє, третій записує. Цей підхід лежить в основі ' +
    'паттерну Producer-Consumer. У java.util.concurrent реалізується через ExecutorService з різними ' +
    'типами Callable/Runnable задач.'
  ));
  
  content.push(heading2('3.3. Конвеєрна обробка (Pipeline Parallelism)'));
  content.push(para(
    'Задача розбивається на послідовні стадії (етапи), кожна стадія виконується паралельно з іншими ' +
    'стадіями на різних даних — аналогія до заводського конвеєра. Поки перший потік обробляє другий ' +
    'елемент, другий потік вже виконує операцію над першим. Забезпечує рівномірне завантаження ресурсів.'
  ));
  
  content.push(heading2('3.4. Рекурсивне розбиття (Divide and Conquer / Fork-Join)'));
  content.push(para(
    'Задача рекурсивно ділиться на менші підзадачі до досягнення базового випадку, після чого ' +
    'результати об\'єднуються. У Java цей підхід реалізується через ForkJoinPool та класи ' +
    'RecursiveTask<V> і RecursiveAction. Алгоритми: merge sort, quicksort, пошук у файловій системі.'
  ));
  content.push(code('// Приклад Fork-Join'));
  content.push(code('class SumTask extends RecursiveTask<Long> {'));
  content.push(code('    long[] data; int lo, hi;'));
  content.push(code('    protected Long compute() {'));
  content.push(code('        if (hi - lo <= THRESHOLD) return sequentialSum(lo, hi);'));
  content.push(code('        int mid = (lo + hi) / 2;'));
  content.push(code('        SumTask left = new SumTask(data, lo, mid);'));
  content.push(code('        left.fork();  // асинхронний запуск'));
  content.push(code('        long rightResult = new SumTask(data, mid, hi).compute();'));
  content.push(code('        return left.join() + rightResult; // об\'єднання'));
  content.push(code('    }'));
  content.push(code('}'));
  
  content.push(heading2('3.5. Розпаралелювання на рівні пулів потоків'));
  content.push(para(
    'FixedThreadPool, CachedThreadPool, ScheduledThreadPool — готові механізми з java.util.concurrent ' +
    'для керування набором робочих потоків. Задачі передаються у чергу (BlockingQueue), потоки ' +
    'беруть їх звідти і виконують. Не потрібно вручну створювати та знищувати потоки, що знижує ' +
    'накладні витрати і спрощує код.'
  ));
  
  content.push(heading2('3.6. Розпаралелювання на рівні даних (MPI, розподілені системи)'));
  content.push(para(
    'При використанні MPI кожен процес отримує свою частину даних і виконує обчислення незалежно. ' +
    'Результати збираються через колективні операції: MPI_Reduce, MPI_Gather, MPI_Bcast. ' +
    'Це найбільш масштабований підхід для суперкомп\'ютерів і кластерів з тисячами вузлів.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 5
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 5. Способи підвищення продуктивності комп\'ютерів. Суперкомп\'ютери'));
  
  content.push(heading2('5.1. Апаратні способи підвищення продуктивності'));
  content.push(para(
    'Протягом десятиліть продуктивність підвищувалась за рахунок збільшення тактової частоти процесора. ' +
    'Однак після ~2004 року зіткнулись із "стіною потужності" (power wall): подальше збільшення частоти ' +
    'неможливе через надмірне тепловиділення. Тому перейшли до багатоядерних архітектур.'
  ));
  content.push(bullet('Збільшення кількості ядер (multi-core): сучасні CPU мають від 4 до 128+ ядер'));
  content.push(bullet('Гіперпоточність (HyperThreading / SMT): одне фізичне ядро виглядає як 2 логічних'));
  content.push(bullet('SIMD (SSE, AVX): обробка кількох елементів даних однією інструкцією'));
  content.push(bullet('GPU (GPGPU): тисячі простих ядер для паралельних обчислень (CUDA, OpenCL)'));
  content.push(bullet('Кеш-пам\'ять L1/L2/L3: зменшення затримок доступу до даних'));
  content.push(bullet('NUMA (Non-Uniform Memory Access): оптимізована топологія пам\'яті у багатопроцесорних системах'));
  
  content.push(heading2('5.2. Програмні способи підвищення продуктивності'));
  content.push(bullet('Оптимізація алгоритмів (зниження складності з O(n²) до O(n log n))'));
  content.push(bullet('Паралельні алгоритми (розбиття задачі між потоками/процесами)'));
  content.push(bullet('Векторизація (автоматична чи ручна оптимізація під SIMD)'));
  content.push(bullet('Профілювання та усунення вузьких місць (Java: JProfiler, VisualVM)'));
  content.push(bullet('Lock-free структури даних (java.util.concurrent.atomic)'));
  content.push(bullet('JIT-компіляція у Java (Just-In-Time): HotSpot компілює гарячий код у машинний'));
  
  content.push(heading2('5.3. Суперкомп\'ютери'));
  content.push(para(
    'Суперкомп\'ютер — це обчислювальна система, яка має продуктивність, що значно перевищує ' +
    'можливості стандартних комп\'ютерів, виміряну у FLOPS (Floating Point Operations Per Second). ' +
    'Сучасні суперкомп\'ютери класу Exascale досягають 10^18 FLOPS.'
  ));
  content.push(para(
    'Архітектура суперкомп\'ютерів: кластери з тисячами вузлів, з\'єднаних мережами InfiniBand або ' +
    'Omni-Path з надмалою латентністю. Кожен вузол може мати кілька багатоядерних CPU та GPU/TPU ' +
    'акселераторів. Програмування здійснюється через MPI + OpenMP (гібридна модель).'
  ));
  content.push(para(
    'Приклади суперкомп\'ютерів: Frontier (Oak Ridge, США, ~1.2 Exaflops), Fugaku (Японія), ' +
    'Summit, Sierra. В Україні є суперкомп\'ютери в Інституті кібернетики НАН України.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 7
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 7. Поняття потоку обчислень. Створення та запуск потоків в Java'));
  
  content.push(heading2('7.1. Що таке потік обчислень'));
  content.push(para(
    'Потік обчислень (thread) — це найменша одиниця виконання в операційній системі, ' +
    'що виконується в межах процесу. Кілька потоків одного процесу поділяють спільну пам\'ять ' +
    '(heap), але мають власний стек, лічильник команд (Program Counter) і набір регістрів. ' +
    'Це дозволяє потокам ефективно обмінюватися даними, але вимагає синхронізації.'
  ));
  content.push(para(
    'Процес vs. Потік: процес — важка одиниця (окремий адресний простір, ресурси ОС), ' +
    'потік — легка одиниця всередині процесу. Створення потоку у 10–100 разів швидше за ' +
    'створення процесу. У JVM кожна програма запускається як процес, і JVM сама створює ' +
    'кілька службових потоків (GC, JIT compiler, finalizer тощо).'
  ));
  
  content.push(heading2('7.2. Способи створення потоків у Java'));
  content.push(bold('Спосіб 1: Розширення класу Thread'));
  content.push(code('class MyThread extends Thread {'));
  content.push(code('    @Override'));
  content.push(code('    public void run() {'));
  content.push(code('        System.out.println("Потік: " + Thread.currentThread().getName());'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code('// Запуск:'));
  content.push(code('MyThread t = new MyThread();'));
  content.push(code('t.start();  // НЕ t.run()! start() створює новий потік'));
  
  content.push(bold('Спосіб 2: Реалізація інтерфейсу Runnable (рекомендований)'));
  content.push(code('class MyTask implements Runnable {'));
  content.push(code('    @Override'));
  content.push(code('    public void run() {'));
  content.push(code('        System.out.println("Runnable потік");'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code('Thread t = new Thread(new MyTask());'));
  content.push(code('t.start();'));
  
  content.push(bold('Спосіб 3: Lambda (Java 8+)'));
  content.push(code('Thread t = new Thread(() -> System.out.println("Lambda потік"));'));
  content.push(code('t.start();'));
  
  content.push(bold('Спосіб 4: Callable + Future (з поверненням результату)'));
  content.push(code('ExecutorService executor = Executors.newFixedThreadPool(4);'));
  content.push(code('Future<Integer> future = executor.submit(() -> {'));
  content.push(code('    return 42; // результат обчислення'));
  content.push(code('});'));
  content.push(code('int result = future.get(); // блокує до завершення'));
  
  content.push(heading2('7.3. Різниця між start() і run()'));
  content.push(para(
    'Виклик t.run() — це звичайний виклик методу в поточному потоці, без жодного паралелізму. ' +
    'Виклик t.start() — реєструє новий потік в JVM і ОС, після чого ОС планує його виконання. ' +
    'Після start() JVM викликає run() у новому потоці. Метод start() можна викликати лише один ' +
    'раз; повторний виклик кидає IllegalThreadStateException.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 9
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 9. Клас Thread'));
  
  content.push(heading2('9.1. Ієрархія та базові властивості'));
  content.push(para(
    'java.lang.Thread реалізує інтерфейс Runnable і є основним класом для роботи з потоками. ' +
    'Кожен об\'єкт Thread має: унікальний ідентифікатор (long id), ім\'я (String name), ' +
    'пріоритет (int priority від Thread.MIN_PRIORITY=1 до Thread.MAX_PRIORITY=10, за замовчуванням NORM_PRIORITY=5), ' +
    'прапор daemon, групу потоків (ThreadGroup), та стан (State).'
  ));
  
  content.push(heading2('9.2. Стани потоку (Thread.State)'));
  content.push(bullet('NEW — створено об\'єкт Thread, але start() ще не викликано'));
  content.push(bullet('RUNNABLE — потік виконується або готовий до виконання (чекає на квант CPU)'));
  content.push(bullet('BLOCKED — чекає на монітор (synchronized блок зайнятий іншим потоком)'));
  content.push(bullet('WAITING — чекає невизначено (після wait(), join() без таймауту)'));
  content.push(bullet('TIMED_WAITING — чекає обмежений час (sleep(ms), wait(ms), join(ms))'));
  content.push(bullet('TERMINATED — метод run() завершив виконання'));
  
  content.push(heading2('9.3. Ключові методи класу Thread'));
  content.push(code('// Статичні методи:'));
  content.push(code('Thread.sleep(1000);          // призупинити поточний потік на 1 с'));
  content.push(code('Thread.yield();              // поступитися CPU (підказка планувальнику)'));
  content.push(code('Thread currentThread = Thread.currentThread(); // посилання на поточний потік'));
  
  content.push(code('// Методи екземпляра:'));
  content.push(code('t.start();                   // запустити потік'));
  content.push(code('t.join();                    // чекати завершення потоку t'));
  content.push(code('t.join(5000);                // чекати не більше 5 секунд'));
  content.push(code('t.interrupt();               // перервати потік (встановлює прапор)'));
  content.push(code('t.isInterrupted();           // перевірити прапор переривання'));
  content.push(code('t.isAlive();                 // чи виконується потік'));
  content.push(code('t.setDaemon(true);           // зробити потік демоном (до start()!)'));
  content.push(code('t.setPriority(Thread.MAX_PRIORITY); // встановити пріоритет'));
  content.push(code('t.getName() / t.setName("worker-1"); // ім\'я потоку'));
  
  content.push(heading2('9.4. Потоки-демони (Daemon Threads)'));
  content.push(para(
    'Daemon-потоки — це службові потоки, що працюють у фоні (наприклад, Garbage Collector). ' +
    'JVM автоматично завершує всі daemon-потоки, як тільки не залишилось жодного не-daemon потоку. ' +
    'Встановлюється через t.setDaemon(true) до виклику t.start(). Корисно для фонових задач ' +
    '(логування, моніторинг), які не повинні заважати завершенню програми.'
  ));
  
  content.push(heading2('9.5. Переривання потоку'));
  content.push(para(
    'Java не має механізму примусового зупинення потоку (Thread.stop() застарілий і небезпечний). ' +
    'Замість цього використовується кооперативний підхід: викликається t.interrupt(), що встановлює ' +
    'прапор переривання. Якщо потік знаходиться у sleep/wait/join, кидається InterruptedException. ' +
    'Якщо потік не заблокований, він має сам перевіряти Thread.currentThread().isInterrupted() і ' +
    'коректно завершуватись.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 11
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 11. Стрічковий алгоритм паралельного множення матриць'));
  
  content.push(heading2('11.1. Задача множення матриць'));
  content.push(para(
    'Класичне множення матриць A(n×k) * B(k×m) = C(n×m): кожен елемент c[i][j] = сума a[i][t]*b[t][j] ' +
    'для t від 0 до k-1. Складність послідовного алгоритму — O(n³). Для великих матриць (1000×1000 і більше) ' +
    'це надзвичайно затратна операція, що ідеально підходить для паралелізації.'
  ));
  
  content.push(heading2('11.2. Ідея стрічкового алгоритму'));
  content.push(para(
    'При стрічковому (strip / row decomposition) підході матриця A ділиться на горизонтальні смуги ' +
    '(strips) між p потоками. Кожен потік i отримує рядки від i*(n/p) до (i+1)*(n/p) матриці A ' +
    'і обчислює відповідні рядки результуючої матриці C. Матриця B при цьому залишається ' +
    'доступною всім потокам (shared, тільки читання — синхронізація не потрібна).'
  ));
  
  content.push(heading2('11.3. Реалізація в Java'));
  content.push(code('int n = 1000, p = Runtime.getRuntime().availableProcessors();'));
  content.push(code('double[][] A = new double[n][n];'));
  content.push(code('double[][] B = new double[n][n];'));
  content.push(code('double[][] C = new double[n][n];'));
  content.push(code(''));
  content.push(code('ExecutorService pool = Executors.newFixedThreadPool(p);'));
  content.push(code('int strip = n / p;'));
  content.push(code('for (int t = 0; t < p; t++) {'));
  content.push(code('    final int rowStart = t * strip;'));
  content.push(code('    final int rowEnd   = (t == p-1) ? n : rowStart + strip;'));
  content.push(code('    pool.submit(() -> {'));
  content.push(code('        for (int i = rowStart; i < rowEnd; i++) {'));
  content.push(code('            for (int j = 0; j < n; j++) {'));
  content.push(code('                double sum = 0;'));
  content.push(code('                for (int k = 0; k < n; k++)'));
  content.push(code('                    sum += A[i][k] * B[k][j];'));
  content.push(code('                C[i][j] = sum;'));
  content.push(code('            }'));
  content.push(code('        }'));
  content.push(code('    });'));
  content.push(code('}'));
  content.push(code('pool.shutdown();'));
  content.push(code('pool.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);'));
  
  content.push(heading2('11.4. Аналіз ефективності'));
  content.push(para(
    'Стрічковий алгоритм має ідеальний баланс навантаження при n % p == 0. Кожен потік виконує ' +
    'рівно n/p рядків. Комунікаційні витрати мінімальні — матриця B читається всіма потоками, ' +
    'але без запису (race condition відсутній). Прискорення наближається до p (кількості потоків), ' +
    'але обмежується законом Амдала і накладними витратами на синхронізацію.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 13
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 13. Алгоритм Кеннона паралельного множення матриць'));
  
  content.push(heading2('13.1. Контекст та обмеження стрічкового алгоритму'));
  content.push(para(
    'Стрічковий алгоритм добре підходить для архітектур зі спільною пам\'яттю (потоки Java). ' +
    'Але в розподілених системах (MPI, кластери) копіювання всієї матриці B на кожен процес ' +
    'неприйнятне через витрати мережі. Алгоритм Кеннона (Cannon\'s Algorithm, 1969) — ' +
    'класичний алгоритм для рівномірно розподіленого множення матриць у 2D-сітці процесів.'
  ));
  
  content.push(heading2('13.2. Ідея алгоритму Кеннона'));
  content.push(para(
    'Алгоритм розподіляє матриці A і B між p×p процесами у вигляді двовимірної сітки. ' +
    'Кожен процес (i, j) отримує блок A[i][j] і B[i][j] розміру (n/p)×(n/p). ' +
    'Алгоритм виконується у три фази:'
  ));
  content.push(bullet('Ініціалізація: циклічний зсув рядків матриці A ліворуч на i позицій, стовпців B вгору на j позицій'));
  content.push(bullet('Ітерація: p кроків, на кожному кроці кожен процес множить свої локальні блоки і накопичує результат'));
  content.push(bullet('Зсув: після кожного кроку блоки A зсуваються ліворуч на 1, блоки B вгору на 1 (циклічно)'));
  
  content.push(heading2('13.3. Переваги Алгоритму Кеннона'));
  content.push(bullet('Кожен процес передає/отримує лише блоки від сусідів — локальна комунікація'));
  content.push(bullet('Обсяг переданих даних: O(n²/p^0.5) замість O(n²) у наївному підході'));
  content.push(bullet('Ідеальний для 2D-торусних мережей міжз\'єднань'));
  content.push(bullet('Обчислювальна складність на процес: O(n³/p)'));
  
  content.push(heading2('13.4. MPI реалізація (концептуально)'));
  content.push(para(
    'В MPI алгоритм реалізується через MPI_Cart_create (2D декартова сітка процесів), ' +
    'MPI_Cart_shift (знаходження сусідів), MPI_Sendrecv (зсув блоків). ' +
    'Кожен процес виконує локальне множення своїх блоків через стандартний O(n³/p)-алгоритм. ' +
    'Алгоритм Фокса (Fox\'s Algorithm) — аналогічний підхід, але використовує broadcast замість shift.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 15
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 15. Управління потоками в Java'));
  
  content.push(heading2('15.1. Планувальник потоків JVM'));
  content.push(para(
    'JVM делегує планування потоків операційній системі. Планувальник ОС вирішує, який потік ' +
    'виконується в кожний момент, з урахуванням пріоритету (setPriority), стану і системної ' +
    'політики. JVM не гарантує певного порядку виконання потоків — це важливо для правильного ' +
    'розуміння паралелізму і необхідності синхронізації.'
  ));
  
  content.push(heading2('15.2. Методи управління'));
  content.push(bold('sleep() — тимчасова пауза'));
  content.push(code('try {'));
  content.push(code('    Thread.sleep(2000); // 2 секунди'));
  content.push(code('} catch (InterruptedException e) {'));
  content.push(code('    Thread.currentThread().interrupt(); // відновити прапор'));
  content.push(code('}'));
  
  content.push(bold('join() — очікування завершення'));
  content.push(code('Thread worker = new Thread(task);'));
  content.push(code('worker.start();'));
  content.push(code('worker.join(); // головний потік чекає завершення worker'));
  
  content.push(bold('yield() — добровільна поступка процесора'));
  content.push(code('Thread.yield(); // підказка планувальнику переключитись'));
  content.push(para(
    'yield() не гарантує фактичного переключення — це лише підказка. Використовується рідко, ' +
    'здебільшого у spin-wait циклах для зменшення навантаження на CPU.'
  ));
  
  content.push(heading2('15.3. Пріоритети потоків'));
  content.push(para(
    'Пріоритети від 1 (MIN_PRIORITY) до 10 (MAX_PRIORITY). Потоки з вищим пріоритетом ' +
    'отримують CPU частіше, але не виключно. Не варто покладатись на пріоритети для забезпечення ' +
    'коректності програми — вони лише впливають на відносну частоту виконання.'
  ));
  
  content.push(heading2('15.4. ThreadGroup'));
  content.push(para(
    'ThreadGroup дозволяє об\'єднувати потоки в логічні групи та керувати ними разом: ' +
    'перервати всі потоки групи (group.interrupt()), отримати активні потоки (group.enumerate()). ' +
    'У сучасному Java рекомендується замість цього використовувати ExecutorService.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 17
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 17. Блокування потоку'));
  
  content.push(heading2('17.1. Поняття блокування'));
  content.push(para(
    'Блокування потоку — це стан, при якому потік призупиняє виконання і очікує настання ' +
    'певної умови: звільнення монітора (synchronized), отримання сигналу (wait/notify), ' +
    'завершення I/O, завершення іншого потоку (join) тощо. Розрізняють блокування ' +
    'BLOCKED (очікування монітора) та WAITING/TIMED_WAITING (очікування умови).'
  ));
  
  content.push(heading2('17.2. Монітори та synchronized'));
  content.push(para(
    'Кожен об\'єкт у Java має вбудований монітор (mutex). Блок synchronized(obj) {} або ' +
    'метод synchronized забезпечує взаємне виключення: лише один потік може утримувати ' +
    'монітор об\'єкта одночасно. Інші потоки, що намагаються увійти у synchronized блок, ' +
    'переходять у стан BLOCKED і чекають.'
  ));
  
  content.push(heading2('17.3. Методи wait(), notify(), notifyAll()'));
  content.push(code('synchronized (lock) {'));
  content.push(code('    while (!condition) {'));
  content.push(code('        lock.wait(); // звільняє монітор і чекає'));
  content.push(code('    }'));
  content.push(code('    // виконати роботу'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// Інший потік:'));
  content.push(code('synchronized (lock) {'));
  content.push(code('    condition = true;'));
  content.push(code('    lock.notifyAll(); // розбудити всіх очікуючих'));
  content.push(code('}'));
  content.push(para(
    'wait() і notify() повинні викликатись лише всередині synchronized блоку, інакше ' +
    'IllegalMonitorStateException. Умова в while (не if!) захищає від spurious wakeups.'
  ));
  
  content.push(heading2('17.4. Явні локери: ReentrantLock'));
  content.push(code('ReentrantLock lock = new ReentrantLock();'));
  content.push(code('Condition condition = lock.newCondition();'));
  content.push(code(''));
  content.push(code('lock.lock();'));
  content.push(code('try {'));
  content.push(code('    while (!ready) condition.await();'));
  content.push(code('    // критична секція'));
  content.push(code('} finally {'));
  content.push(code('    lock.unlock(); // завжди у finally!'));
  content.push(code('}'));
  content.push(para(
    'ReentrantLock надає більше можливостей ніж synchronized: tryLock(timeout), ' +
    'lockInterruptibly(), кілька Condition об\'єктів для одного lock, fairness-режим.'
  ));
  
  content.push(heading2('17.5. Дедлок (Deadlock)'));
  content.push(para(
    'Дедлок виникає, коли два і більше потоки чекають один одного у циклічній залежності. ' +
    'Потік A тримає ресурс X і чекає на Y, потік B тримає Y і чекає на X. Жоден не може ' +
    'продовжити. Запобігання: завжди захоплювати локери в одному і тому ж порядку; ' +
    'використовувати tryLock з таймаутом; уникати вкладених synchronized блоків.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 19
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 19. Синхронізовані методи'));
  
  content.push(heading2('19.1. Ключове слово synchronized'));
  content.push(para(
    'Ключове слово synchronized у Java забезпечує атомарність виконання коду та видимість змін ' +
    'між потоками (happens-before гарантія). Використовується для захисту shared state від ' +
    'race conditions — ситуацій, коли результат залежить від порядку виконання потоків.'
  ));
  
  content.push(heading2('19.2. Синхронізовані методи екземпляра'));
  content.push(code('class Counter {'));
  content.push(code('    private int count = 0;'));
  content.push(code('    '));
  content.push(code('    public synchronized void increment() {'));
  content.push(code('        count++;  // атомарна операція'));
  content.push(code('    }'));
  content.push(code('    '));
  content.push(code('    public synchronized int getCount() {'));
  content.push(code('        return count;'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(para(
    'Монітором є this — поточний об\'єкт. Лише один потік може одночасно виконувати ' +
    'будь-який synchronized метод цього об\'єкта.'
  ));
  
  content.push(heading2('19.3. Синхронізовані статичні методи'));
  content.push(code('class SharedCounter {'));
  content.push(code('    private static int count = 0;'));
  content.push(code('    '));
  content.push(code('    public static synchronized void increment() {'));
  content.push(code('        count++;'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(para(
    'Монітором є об\'єкт Class (SharedCounter.class). Статичні synchronized методи ' +
    'синхронізуються між усіма екземплярами класу.'
  ));
  
  content.push(heading2('19.4. Synchronized блоки'));
  content.push(code('class DataProcessor {'));
  content.push(code('    private final Object lock = new Object();'));
  content.push(code('    private List<String> data = new ArrayList<>();'));
  content.push(code('    '));
  content.push(code('    public void addItem(String item) {'));
  content.push(code('        synchronized (lock) {  // блокуємо лише критичну секцію'));
  content.push(code('            data.add(item);'));
  content.push(code('        }'));
  content.push(code('        // решта коду виконується без блокування'));
  content.push(code('    }'));
  content.push(code('}'));
  
  content.push(heading2('19.5. Видимість змін: ключове слово volatile'));
  content.push(code('class FlagExample {'));
  content.push(code('    private volatile boolean running = true;'));
  content.push(code('    '));
  content.push(code('    public void stop() { running = false; }'));
  content.push(code('    '));
  content.push(code('    public void run() {'));
  content.push(code('        while (running) {  // потік побачить оновлення без synchronized'));
  content.push(code('            doWork();'));
  content.push(code('        }'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(para(
    'volatile гарантує видимість, але не атомарність. Підходить для прапорців. ' +
    'Для атомарних операцій краще використовувати AtomicInteger, AtomicBoolean тощо.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 21
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 21. Високорівневі способи управління потоками з java.util.concurrent'));
  
  content.push(heading2('21.1. Огляд пакету java.util.concurrent'));
  content.push(para(
    'Пакет java.util.concurrent (JUC), введений у Java 5, надає готові рішення для ' +
    'багатопоточного програмування: пули потоків, черги, синхронізатори, атомарні змінні, ' +
    'concurrent колекції. Він суттєво спрощує написання коректних та ефективних паралельних ' +
    'програм порівняно з низькорівневим Thread/synchronized.'
  ));
  
  content.push(heading2('21.2. ExecutorService та Executors'));
  content.push(code('// Фабричні методи Executors:'));
  content.push(code('ExecutorService fixed   = Executors.newFixedThreadPool(4);'));
  content.push(code('ExecutorService cached  = Executors.newCachedThreadPool();'));
  content.push(code('ExecutorService single  = Executors.newSingleThreadExecutor();'));
  content.push(code('ScheduledExecutorService sched = Executors.newScheduledThreadPool(2);'));
  content.push(para(
    'ExecutorService — інтерфейс для управління пулами потоків. Метод submit() повертає ' +
    'Future<T>, через який можна отримати результат або обробити виняток. shutdown() ' +
    'ініціює коректне завершення (нові задачі не приймаються, поточні доробляються), ' +
    'shutdownNow() намагається перервати виконувані задачі.'
  ));
  
  content.push(heading2('21.3. Future та CompletableFuture'));
  content.push(code('Future<Double> future = executor.submit(() -> computePi());'));
  content.push(code('double pi = future.get(5, TimeUnit.SECONDS); // з таймаутом'));
  content.push(para(
    'CompletableFuture (Java 8+) дозволяє будувати асинхронні конвеєри:'
  ));
  content.push(code('CompletableFuture.supplyAsync(() -> fetchData())'));
  content.push(code('    .thenApply(data -> process(data))'));
  content.push(code('    .thenAccept(result -> save(result))'));
  content.push(code('    .exceptionally(e -> { log(e); return null; });'));
  
  content.push(heading2('21.4. Синхронізатори'));
  content.push(bullet('CountDownLatch: дозволяє одному/кільком потокам чекати, поки N інших завершать операцію'));
  content.push(bullet('CyclicBarrier: N потоків чекають один одного на бар\'єрній точці (використовується повторно)'));
  content.push(bullet('Semaphore: обмежує кількість одночасних доступів до ресурсу (наприклад, до БД)'));
  content.push(bullet('Exchanger: двоє потоків обмінюються об\'єктами у точці зустрічі'));
  content.push(bullet('Phaser: гнучкіша версія CyclicBarrier з динамічним числом учасників'));
  
  content.push(heading2('21.5. Concurrent колекції'));
  content.push(bullet('ConcurrentHashMap: потокобезпечна HashMap без глобального блокування'));
  content.push(bullet('CopyOnWriteArrayList: при записі створюється копія, ідеально для частого читання'));
  content.push(bullet('ArrayBlockingQueue / LinkedBlockingQueue: потокобезпечні черги для Producer-Consumer'));
  content.push(bullet('PriorityBlockingQueue: черга з пріоритетом'));
  content.push(bullet('ConcurrentLinkedQueue: lock-free черга на основі CAS-операцій'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 23
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 23. Пули потоків'));
  
  content.push(heading2('23.1. Навіщо потрібні пули потоків'));
  content.push(para(
    'Створення і знищення потоку — дорога операція: виділення пам\'яті для стеку (~512KB за замовч.), ' +
    'системний виклик ОС, ініціалізація структур. При обробці тисяч коротких задач накладні витрати ' +
    'перевищують корисну роботу. Пул потоків (thread pool) вирішує проблему: завчасно створюються ' +
    'N потоків, що багаторазово беруть задачі з черги й виконують їх.'
  ));
  
  content.push(heading2('23.2. ThreadPoolExecutor — основа всіх пулів'));
  content.push(code('ThreadPoolExecutor executor = new ThreadPoolExecutor('));
  content.push(code('    4,          // corePoolSize: мінімальна кількість потоків'));
  content.push(code('    8,          // maximumPoolSize: максимальна кількість потоків'));
  content.push(code('    60, TimeUnit.SECONDS, // keepAliveTime для зайвих потоків'));
  content.push(code('    new LinkedBlockingQueue<>(100), // черга задач'));
  content.push(code('    new ThreadFactory() { ... },    // фабрика потоків'));
  content.push(code('    new ThreadPoolExecutor.CallerRunsPolicy() // RejectionHandler'));
  content.push(code(');'));
  
  content.push(heading2('23.3. FixedThreadPool'));
  content.push(para(
    'Executors.newFixedThreadPool(n) — пул із фіксованою кількістю потоків. ' +
    'Підходить для CPU-bound задач, де оптимальна кількість потоків ≈ кількість ядер. ' +
    'Черга необмежена (LinkedBlockingQueue), тому задачі накопичуються при навантаженні. ' +
    'Ідеально для систем масового обслуговування (СМО), де є фіксована кількість "серверів".'
  ));
  content.push(code('ExecutorService executor = Executors.newFixedThreadPool('));
  content.push(code('    Runtime.getRuntime().availableProcessors()'));
  content.push(code(');'));
  content.push(code('for (int i = 0; i < 1000; i++) {'));
  content.push(code('    executor.submit(() -> processRequest());'));
  content.push(code('}'));
  content.push(code('executor.shutdown();'));
  
  content.push(heading2('23.4. CachedThreadPool'));
  content.push(para(
    'Executors.newCachedThreadPool() — пул із необмеженою кількістю потоків. ' +
    'Потоки, що простоювали 60 секунд, видаляються. Нові потоки створюються за потребою. ' +
    'Підходить для I/O-bound задач з коротким часом виконання. Небезпечний при великій ' +
    'кількості задач — може породити тисячі потоків і вичерпати пам\'ять.'
  ));
  
  content.push(heading2('23.5. ScheduledThreadPool'));
  content.push(para(
    'Executors.newScheduledThreadPool(n) — пул для відкладеного та повторюваного виконання задач.'
  ));
  content.push(code('ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(2);'));
  content.push(code('scheduler.scheduleAtFixedRate(() -> checkMetrics(), 0, 5, TimeUnit.SECONDS);'));
  content.push(code('scheduler.schedule(() -> sendReport(), 30, TimeUnit.MINUTES);'));
  
  content.push(heading2('23.6. Стратегії відхилення (RejectionHandler)'));
  content.push(bullet('AbortPolicy (за замовч.): кидає RejectedExecutionException'));
  content.push(bullet('CallerRunsPolicy: задача виконується в потоці, що викликав submit()'));
  content.push(bullet('DiscardPolicy: задача мовчки відкидається'));
  content.push(bullet('DiscardOldestPolicy: видаляється найстарша задача з черги'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 25
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 25. Розробка паралельних програм з використанням ForkJoinFramework'));
  
  content.push(heading2('25.1. Концепція Fork-Join'));
  content.push(para(
    'ForkJoinPool (Java 7+) реалізує парадигму "розділяй і влаштовуй" (divide & conquer). ' +
    'Задача рекурсивно ділиться на підзадачі (fork), кожна з яких може знову ділитись. ' +
    'Результати рекурсивно об\'єднуються (join). Ключова інновація — work stealing: ' +
    'якщо потік виконав свої задачі, він "краде" задачі з черги іншого зайнятого потоку, ' +
    'забезпечуючи рівномірне завантаження.'
  ));
  
  content.push(heading2('25.2. Ключові класи'));
  content.push(bullet('ForkJoinPool: спеціальний пул для Fork-Join задач (є загальний common pool)'));
  content.push(bullet('RecursiveTask<V>: задача, що повертає результат (потрібно перевизначити compute())'));
  content.push(bullet('RecursiveAction: задача без повертання результату (наприклад, сортування на місці)'));
  content.push(bullet('ForkJoinTask<V>: базовий клас для обох'));
  
  content.push(heading2('25.3. Приклад: пошук файлів у директорії'));
  content.push(code('class FileSearchTask extends RecursiveTask<List<Path>> {'));
  content.push(code('    private final Path dir;'));
  content.push(code('    private final String pattern;'));
  content.push(code('    '));
  content.push(code('    @Override'));
  content.push(code('    protected List<Path> compute() {'));
  content.push(code('        List<Path> results = new ArrayList<>();'));
  content.push(code('        List<FileSearchTask> subTasks = new ArrayList<>();'));
  content.push(code('        '));
  content.push(code('        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir)) {'));
  content.push(code('            for (Path entry : stream) {'));
  content.push(code('                if (Files.isDirectory(entry)) {'));
  content.push(code('                    FileSearchTask subTask = new FileSearchTask(entry, pattern);'));
  content.push(code('                    subTask.fork();  // асинхронний запуск підзадачі'));
  content.push(code('                    subTasks.add(subTask);'));
  content.push(code('                } else if (entry.toString().contains(pattern)) {'));
  content.push(code('                    results.add(entry);'));
  content.push(code('                }'));
  content.push(code('            }'));
  content.push(code('        }'));
  content.push(code('        for (FileSearchTask sub : subTasks) {'));
  content.push(code('            results.addAll(sub.join()); // збираємо результати'));
  content.push(code('        }'));
  content.push(code('        return results;'));
  content.push(code('    }'));
  content.push(code('}'));
  content.push(code(''));
  content.push(code('// Запуск:'));
  content.push(code('ForkJoinPool pool = new ForkJoinPool();'));
  content.push(code('List<Path> found = pool.invoke(new FileSearchTask(Paths.get("/"), ".java"));'));
  
  content.push(heading2('25.4. Work Stealing'));
  content.push(para(
    'Кожен потік у ForkJoinPool має власну двосторонню чергу (deque) задач. ' +
    'Потік додає і бере задачі зі свого кінця (LIFO — ефективно для кешу). ' +
    'Якщо потік простоює, він "краде" задачі з іншого кінця чужої черги (FIFO). ' +
    'Це мінімізує синхронізацію і забезпечує автоматичний балансування навантаження.'
  ));
  
  content.push(heading2('25.5. Поради по використанню'));
  content.push(bullet('Встановлюйте threshold: при малому розмірі задачі виконуйте її послідовно (уникайте overhead)'));
  content.push(bullet('Використовуйте invokeAll() для розгалуження симетричних підзадач'));
  content.push(bullet('ForkJoinPool.commonPool() — глобальний пул, доступний без явного створення'));
  content.push(bullet('Уникайте I/O та блокуючих операцій всередині RecursiveTask'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 27
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 27. Модель обчислень у вигляді графа «операції—операнди»'));
  
  content.push(heading2('27.1. Визначення моделі'));
  content.push(para(
    'Граф операції-операнди (або граф обчислень, computation graph) — це орієнтований ациклічний ' +
    'граф (DAG), де вершини — це або операції (обчислення), або операнди (дані/змінні), ' +
    'а дуги показують залежності: яка операція породжує який операнд, який операнд споживається ' +
    'якою операцією. Ця модель дозволяє аналізувати паралелізм задачі незалежно від конкретної ' +
    'архітектури.'
  ));
  
  content.push(heading2('27.2. Аналіз залежностей'));
  content.push(para(
    'Дві операції можуть виконуватись паралельно тоді і тільки тоді, коли між ними немає шляху ' +
    'у графі залежностей (тобто жодна не залежить від результатів іншої). ' +
    'Граф виявляє: справжній паралелізм (true parallelism), антизалежності (write after read), ' +
    'залежності по виводу (write after write).'
  ));
  
  content.push(heading2('27.3. Критичний шлях'));
  content.push(para(
    'Критичний шлях (critical path) — найдовший шлях від входу до виходу в графі залежностей. ' +
    'Довжина критичного шляху визначає теоретичний нижній поріг часу виконання незалежно від ' +
    'кількості процесорів. Ступінь паралелізму = загальна робота / довжина критичного шляху. ' +
    'Це ключовий показник для оцінки потенційного прискорення.'
  ));
  
  content.push(heading2('27.4. Застосування'));
  content.push(bullet('Компілятори: визначення порядку виконання інструкцій для SIMD/суперскалярних процесорів'));
  content.push(bullet('MapReduce/Spark: DAG задач, кожна стадія може виконуватись паралельно'));
  content.push(bullet('Build systems (Make, Gradle): паралельна збірка незалежних модулів'));
  content.push(bullet('Нейронні мережі: TensorFlow/PyTorch будують граф обчислень для авто-диференціювання'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 29
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 29. Показники ефективності паралельного алгоритму'));
  
  content.push(heading2('29.1. Час виконання'));
  content.push(para(
    'T₁ — час виконання на одному процесорі (послідовний час). ' +
    'Tₚ — час виконання на p процесорах. ' +
    'T∞ — час виконання на нескінченній кількості процесорів (= довжина критичного шляху). ' +
    'Всі ці величини вимірюються в секундах або у числі операцій.'
  ));
  
  content.push(heading2('29.2. Прискорення (Speedup)'));
  content.push(para(
    'Прискорення Sp = T₁ / Tₚ показує, у скільки разів паралельний алгоритм швидший за послідовний. ' +
    'Лінійне прискорення: Sp = p (ідеально). Суперлінійне прискорення: Sp > p (буває при кращому ' +
    'використанні кешу). Субліне прискорення: Sp < p (реальна ситуація через накладні витрати).'
  ));
  
  content.push(heading2('29.3. Ефективність (Efficiency)'));
  content.push(para(
    'Ефективність Ep = Sp / p = T₁ / (p * Tₚ) — міра корисного використання кожного процесора. ' +
    'Ep = 1 (100%) — ідеально, кожен процесор завантажений весь час. ' +
    'Ep < 1 — є накладні витрати (комунікація, синхронізація, незбалансованість навантаження). ' +
    'Зазвичай ефективність знижується при збільшенні p.'
  ));
  
  content.push(heading2('29.4. Масштабованість (Scalability)'));
  content.push(para(
    'Сильна масштабованість: при фіксованому розмірі задачі збільшення p дає пропорційне прискорення. ' +
    'Слабка масштабованість (закон Густафсона): при збільшенні p пропорційно збільшується розмір задачі, ' +
    'ефективність залишається постійною. Слабка масштабованість більш реалістична для великих систем.'
  ));
  
  content.push(heading2('29.5. Overhead паралелізму'));
  content.push(bullet('Комунікаційний overhead: передача даних між процесами/потоками'));
  content.push(bullet('Синхронізаційний overhead: очікування на бар\'єрах, mutexах'));
  content.push(bullet('Load imbalance: нерівномірний розподіл роботи між процесорами'));
  content.push(bullet('Overhead запуску: час на створення потоків, ініціалізацію'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 31
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 31. Закон Амдала. Ефект Амдала'));
  
  content.push(heading2('31.1. Формулювання закону'));
  content.push(para(
    'Закон Амдала (Gene Amdahl, 1967) встановлює теоретичну верхню межу прискорення паралельної ' +
    'програми. Нехай f — частка програми, що не може бути розпаралелена (послідовна частина, ' +
    '0 ≤ f ≤ 1), а (1-f) — частка, що розпаралелюється. Тоді максимальне прискорення ' +
    'при p процесорах:'
  ));
  content.push(para(
    'S(p) = 1 / (f + (1-f)/p)'
  ));
  content.push(para(
    'При p → ∞: S(∞) = 1/f. Тобто навіть нескінченна кількість процесорів не може подолати ' +
    'послідовну частину. При f = 0.1 (10% послідовного): максимальне прискорення = 10x. ' +
    'При f = 0.05: максимум = 20x.'
  ));
  
  content.push(heading2('31.2. Ефект Амдала'));
  content.push(para(
    'Ефект Амдала — це практичний висновок: навіть невелика частка послідовного коду жорстко ' +
    'обмежує паралельну продуктивність. Якщо 20% коду не піддається розпаралелюванню, ' +
    'максимальне прискорення не перевищить 5x, незалежно від кількості процесорів. ' +
    'Це стимулює до мінімізації послідовних секцій і критичних шляхів.'
  ));
  
  content.push(heading2('31.3. Закон Густафсона-Барсіса (контраргумент)'));
  content.push(para(
    'Густафсон (1988) запропонував альтернативний погляд: при збільшенні p збільшується ' +
    'і розмір задачі, тому абсолютна тривалість паралельної частини зростає, а послідовна — ні. ' +
    'Scaled speedup = p - f(p-1). Це пояснює, чому реальні системи добре масштабуються: ' +
    'більше процесорів дозволяє вирішувати більші задачі за той же час.'
  ));
  
  content.push(heading2('31.4. Приклад обчислення'));
  content.push(code('// Нехай f = 0.2 (20% послідовного коду)'));
  content.push(code('// p = 4: S = 1/(0.2 + 0.8/4) = 1/(0.2 + 0.2) = 1/0.4 = 2.5x'));
  content.push(code('// p = 8: S = 1/(0.2 + 0.8/8) = 1/(0.2 + 0.1) = 1/0.3 = 3.33x'));
  content.push(code('// p = ∞: S = 1/0.2 = 5x (максимум!)'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 33
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 33. Оцінка ефективності паралельних обчислень'));
  
  content.push(heading2('33.1. Теоретичні методи оцінки'));
  content.push(para(
    'Теоретичну оцінку будують на основі аналізу алгоритму: рахують загальний обсяг роботи T₁, ' +
    'визначають критичний шлях T∞, обчислюють ступінь паралелізму T₁/T∞. ' +
    'Верхня межа прискорення: min(p, T₁/T∞). Обчислення роботи: для множення матриць n×n — ' +
    'T₁ = O(n³), T∞ = O(log n) при необмеженій паралелізації.'
  ));
  
  content.push(heading2('33.2. Практичні методи вимірювання'));
  content.push(bullet('Профілювання (Java: JMH, JProfiler, VisualVM): вимірювання реального часу виконання'));
  content.push(bullet('Benchmarking: запуск при різних p і вимірювання Tₚ'));
  content.push(bullet('Hardware Performance Counters (PAPI): вимірювання cache misses, IPC'));
  content.push(bullet('Аналіз trace-файлів: запис подій кожного потоку і побудова діаграм Ганта'));
  
  content.push(heading2('33.3. Мікробенчмаркінг в Java (JMH)'));
  content.push(code('@Benchmark'));
  content.push(code('@BenchmarkMode(Mode.AverageTime)'));
  content.push(code('@OutputTimeUnit(TimeUnit.MILLISECONDS)'));
  content.push(code('public void matrixMultiply(Blackhole bh) {'));
  content.push(code('    double[][] result = parallelMultiply(A, B, nThreads);'));
  content.push(code('    bh.consume(result); // запобігає оптимізації компілятора'));
  content.push(code('}'));
  
  content.push(heading2('33.4. Аналіз результатів'));
  content.push(para(
    'Будуємо графіки: Tₚ від p (крива прискорення) та Ep від p (крива ефективності). ' +
    'Ідеальне прискорення — пряма лінія. Відхилення показують де overhead. ' +
    'Knee of the curve — точка, після якої додавання процесорів майже не дає прискорення: ' +
    'оптимальна кількість процесорів для даної задачі.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 35
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 35. Архітектура розподілених систем. Таксономія Фліна'));
  
  content.push(heading2('35.1. Таксономія Фліна (Flynn\'s Taxonomy, 1966)'));
  content.push(para(
    'Майкл Флін запропонував класифікацію комп\'ютерних архітектур за двома вимірами: ' +
    'потік інструкцій (Single/Multiple Instruction) та потік даних (Single/Multiple Data).'
  ));
  content.push(bullet('SISD (Single Instruction, Single Data): класичний однопроцесорний комп\'ютер Фон Неймана'));
  content.push(bullet('SIMD (Single Instruction, Multiple Data): векторні процесори, GPU, MMX/SSE/AVX у CPU'));
  content.push(bullet('MISD (Multiple Instruction, Single Data): теоретична, рідко зустрічається (fault-tolerant системи)'));
  content.push(bullet('MIMD (Multiple Instruction, Multiple Data): сучасні multi-core CPU, кластери, Grid — найпоширеніша категорія'));
  
  content.push(heading2('35.2. MIMD підкатегорії'));
  content.push(bullet('Shared Memory MIMD (SMP, NUMA): процесори мають доступ до спільної пам\'яті; синхронізація через mutex/semaphore'));
  content.push(bullet('Distributed Memory MIMD: кожен процесор має власну пам\'ять; комунікація через повідомлення (MPI)'));
  content.push(bullet('Hybrid: комбінація — вузли зі спільною пам\'яттю, з\'єднані мережею (сучасні кластери: MPI + OpenMP)'));
  
  content.push(heading2('35.3. Архітектура розподілених систем'));
  content.push(para(
    'Розподілена система — це набір незалежних комп\'ютерів, що виглядають для користувача як ' +
    'єдина система. Основні архітектурні моделі:'
  ));
  content.push(bullet('Клієнт-сервер: виділені сервери обробляють запити клієнтів'));
  content.push(bullet('P2P (Peer-to-Peer): рівноправні вузли, кожен є і клієнтом, і сервером'));
  content.push(bullet('Мікросервіси: додаток як набір незалежних сервісів, що спілкуються через API'));
  content.push(bullet('Grid Computing: об\'єднання ресурсів різних організацій для великих обчислень'));
  content.push(bullet('Cloud Computing: надання обчислювальних ресурсів як сервісу (IaaS, PaaS, SaaS)'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 37
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 37. Стандарти та технології обміну повідомленнями в розподілених системах'));
  
  content.push(heading2('37.1. MPI (Message Passing Interface)'));
  content.push(para(
    'MPI — де-факто стандарт для паралельних обчислень на кластерах і суперкомп\'ютерах. ' +
    'Визначає API для відправки та отримання повідомлень між процесами. ' +
    'Реалізації: OpenMPI, MPICH, Intel MPI. MPI процеси мають унікальний ранг (rank) ' +
    'у комунікаторі та обмінюються типізованими повідомленнями.'
  ));
  
  content.push(heading2('37.2. RPC та gRPC'));
  content.push(para(
    'RPC (Remote Procedure Call) — виклик функції на іншому комп\'ютері, що виглядає як ' +
    'локальний виклик. gRPC (Google) — сучасна реалізація на основі HTTP/2 і Protocol Buffers. ' +
    'Підходить для мікросервісних архітектур, забезпечує типізацію, стримінг, мови-агностичний API.'
  ));
  
  content.push(heading2('37.3. Message Brokers'));
  content.push(para(
    'Apache Kafka, RabbitMQ, ActiveMQ — брокери повідомлень для асинхронної комунікації. ' +
    'Виробники (producers) відправляють повідомлення в теми/черги, споживачі (consumers) ' +
    'читають незалежно. Забезпечує відв\'язку компонентів, буферизацію при різній швидкості, ' +
    'гарантії доставки.'
  ));
  
  content.push(heading2('37.4. REST та WebSocket'));
  content.push(para(
    'REST (HTTP) — простий стандарт для мікросервісів, зрозумілий і широко підтримуваний. ' +
    'WebSocket — двонаправлений канал для real-time комунікації. ' +
    'Spring Boot у Java надає готові інструменти для REST API (@RestController) і WebSocket.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 39
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 39. Методи обміну повідомленнями MPI один до одного'));
  
  content.push(heading2('39.1. Блокуючі методи (Blocking Point-to-Point)'));
  content.push(para(
    'Блокуючі методи не повертають управління до завершення операції. MPI_Send блокує ' +
    'відправника до тих пір, поки повідомлення не буде або буферизовано або отримано. ' +
    'MPI_Recv блокує до отримання повідомлення. Прості у використанні, але можуть ' +
    'призвести до дедлоку при неправильному порядку відправки/отримання.'
  ));
  content.push(code('// Процес 0 відправляє, процес 1 отримує:'));
  content.push(code('if (rank == 0) {'));
  content.push(code('    MPI.COMM_WORLD.Send(data, 0, data.length, MPI.DOUBLE, 1, 0);'));
  content.push(code('} else if (rank == 1) {'));
  content.push(code('    MPI.COMM_WORLD.Recv(buf, 0, buf.length, MPI.DOUBLE, 0, 0);'));
  content.push(code('}'));
  
  content.push(heading2('39.2. Неблокуючі методи (Non-blocking)'));
  content.push(para(
    'Isend / Irecv повертають одразу об\'єкт Request, не чекаючи завершення. ' +
    'Процес може продовжувати обчислення, поки відбувається передача. ' +
    'Необхідно викликати Wait() або Test() для перевірки завершення і безпечного доступу до даних. ' +
    'Дозволяє перекрити (overlap) комунікацію та обчислення.'
  ));
  content.push(code('Request req = MPI.COMM_WORLD.Isend(data, 0, n, MPI.DOUBLE, dest, tag);'));
  content.push(code('// ... виконуємо обчислення поки дані передаються ...'));
  content.push(code('req.Wait(); // чекаємо завершення передачі'));
  
  content.push(heading2('39.3. MPI_Sendrecv'));
  content.push(para(
    'MPI_Sendrecv — атомарна операція відправки та отримання. Вирішує проблему дедлоку ' +
    'при кільцевому обміні (наприклад, у стрічковому алгоритмі або алгоритмі Кеннона, ' +
    'де кожен процес відправляє сусіду зліва і отримує від сусіда справа).'
  ));
  content.push(code('MPI.COMM_WORLD.Sendrecv('));
  content.push(code('    sendBuf, 0, n, MPI.DOUBLE, rightNeighbor, 0, // відправка'));
  content.push(code('    recvBuf, 0, n, MPI.DOUBLE, leftNeighbor,  0, // отримання'));
  content.push(code('    MPI.STATUS_IGNORE'));
  content.push(code(');'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 41
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 41. OpenMPI: особливості побудови програми та основні методи'));
  
  content.push(heading2('41.1. Структура MPI програми'));
  content.push(para(
    'Будь-яка MPI програма починається з MPI_Init та закінчується MPI_Finalize. ' +
    'Всі процеси виконують один і той самий код, але розгалужуються за рангом (rank). ' +
    'Комунікатор MPI_COMM_WORLD містить всі запущені процеси. ' +
    'Кількість процесів визначається при запуску: mpirun -np 8 ./program'
  ));
  content.push(code('// Java MPI (mpi4java або через JNI):'));
  content.push(code('MPI.Init(args);'));
  content.push(code('int rank = MPI.COMM_WORLD.Rank();   // ранг процесу (0..p-1)'));
  content.push(code('int size = MPI.COMM_WORLD.Size();   // загальна кількість процесів'));
  content.push(code('// ... логіка програми ...'));
  content.push(code('MPI.Finalize();'));
  
  content.push(heading2('41.2. Колективні операції'));
  content.push(bullet('MPI_Bcast: один процес (root) розсилає дані всім іншим'));
  content.push(bullet('MPI_Scatter: root розбиває масив і роздає частини всім'));
  content.push(bullet('MPI_Gather: збирає частини від всіх процесів до root'));
  content.push(bullet('MPI_Allgather: кожен отримує повний зібраний масив'));
  content.push(bullet('MPI_Reduce: редукція (сума, max, min тощо) до root процесу'));
  content.push(bullet('MPI_Allreduce: редукція з розсилкою результату всім'));
  content.push(bullet('MPI_Barrier: синхронізаційний бар\'єр — всі чекають один одного'));
  
  content.push(heading2('41.3. Типи даних MPI'));
  content.push(para(
    'MPI має власну систему типів: MPI_INT, MPI_DOUBLE, MPI_CHAR та ін., ' +
    'що дозволяє коректно серіалізувати та передавати дані між різними архітектурами ' +
    '(little-endian/big-endian, різні розміри int). Можна визначати власні складові типи ' +
    'через MPI_Type_create_struct.'
  ));
  
  content.push(heading2('41.4. Приклад: паралельне обчислення суми'));
  content.push(code('// Кожен процес рахує часткову суму свого шматка масиву'));
  content.push(code('double localSum = 0;'));
  content.push(code('for (int i = rank * chunkSize; i < (rank+1) * chunkSize; i++)'));
  content.push(code('    localSum += data[i];'));
  content.push(code(''));
  content.push(code('// Редукція: сума localSum всіх процесів -> globalSum у root=0'));
  content.push(code('double[] globalSum = new double[1];'));
  content.push(code('MPI.COMM_WORLD.Reduce(new double[]{localSum}, 0, globalSum, 0,'));
  content.push(code('                     1, MPI.DOUBLE, MPI.SUM, 0);'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 43
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 43. Алгоритми паралельного сортування'));
  
  content.push(heading2('43.1. Паралельне злиття (Parallel Merge Sort)'));
  content.push(para(
    'Масив ділиться між p процесами. Кожен процес сортує свою частину послідовно. ' +
    'Далі виконується дерево злиттів: пари процесів обмінюються частинами і зливають їх. ' +
    'Глибина дерева: log₂(p). Складність: O(n/p * log(n/p) + n/p * log(p)). ' +
    'У Java: ForkJoinPool + Arrays.sort() для підмасивів, потім merge.'
  ));
  
  content.push(heading2('43.2. Бітонне сортування (Bitonic Sort)'));
  content.push(para(
    'Бітонне сортування — паралельний алгоритм для n = 2^k елементів. ' +
    'Побудований на мережах сортування (sorting networks) — фіксованій схемі порівнянь. ' +
    'Глибина паралельного алгоритму: O(log²n). Ідеально для GPU (SIMD) та FPGA. ' +
    'Використовується в CUDA для сортування на GPU.'
  ));
  
  content.push(heading2('43.3. Порозрядне сортування (Parallel Radix Sort)'));
  content.push(para(
    'Паралельний radix sort ділить елементи між процесами за розрядами. ' +
    'Ефективний для цілих чисел, O(n/p * k) де k — кількість розрядів. ' +
    'Активно використовується у GPU-обчисленнях і MapReduce.'
  ));
  
  content.push(heading2('43.4. Sample Sort'));
  content.push(para(
    'Алгоритм для p процесів у MPI: кожен процес сортує локальні дані, обирає p-1 "зразків". ' +
    'Зразки збираються і сортуються для визначення p-1 глобальних сплітерів (splitters). ' +
    'Дані перерозподіляються між процесами згідно сплітерів, кожен процес зливає своїх. ' +
    'Забезпечує рівномірний розподіл незалежно від розподілу вхідних даних.'
  ));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 45
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 45. Оцінювання продуктивності паралельних алгоритмів у розподілених системах'));
  
  content.push(heading2('45.1. Специфіка розподілених систем'));
  content.push(para(
    'На відміну від shared-memory систем, у розподілених системах критичний фактор — ' +
    'комунікаційна затримка (latency) та пропускна здатність мережі (bandwidth). ' +
    'Модель BSP (Bulk Synchronous Parallel) та LogP-модель описують витрати на комунікацію ' +
    'і дозволяють теоретично оцінити час виконання паралельного алгоритму.'
  ));
  
  content.push(heading2('45.2. Метрики для розподілених систем'));
  content.push(bullet('Latency (ms): час на відправку одного повідомлення (overhead + RTT/2)'));
  content.push(bullet('Bandwidth (GB/s): максимальна пропускна здатність каналу'));
  content.push(bullet('Communication-to-computation ratio: відношення часу комунікації до обчислення'));
  content.push(bullet('Прискорення та ефективність (ті ж формули, що й для shared-memory)'));
  content.push(bullet('Ізоефективність: функція w(p), яка визначає наскільки треба збільшити задачу при збільшенні p для збереження ефективності'));
  
  content.push(heading2('45.3. Практичні інструменти'));
  content.push(bullet('MPI profiling: MPE, TAU, Vampir — візуалізація комунікацій у вигляді діаграм Ганта'));
  content.push(bullet('Intel VTune Profiler, Perf: аналіз cache misses, IPC на рівні процесора'));
  content.push(bullet('Java: JMH для мікробенчмарків, VisualVM для heap/thread аналізу'));
  content.push(bullet('LINPACK benchmark: стандартний тест для рейтингу TOP500 суперкомп\'ютерів'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 47
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 47. Грід-технології. Поняття спільного віртуального ресурсу'));
  
  content.push(heading2('47.1. Що таке Grid Computing'));
  content.push(para(
    'Grid Computing (грід-обчислення) — це інфраструктура для координованого використання ' +
    'розподілених гетерогенних ресурсів різних організацій для вирішення великих наукових ' +
    'і технічних задач. Термін запроваджений Яном Фостером та Карлом Кессельманом у 1990-х. ' +
    'Аналогія з електричною мережею (grid): ресурси доступні з "розетки" без знання їх джерела.'
  ));
  
  content.push(heading2('47.2. Спільний віртуальний ресурс (Virtual Organization)'));
  content.push(para(
    'Virtual Organization (VO) — це динамічна група людей і організацій, що поділяють ' +
    'ресурси (процесорний час, сховища, дані, інструменти) для досягнення спільної мети. ' +
    'VO не прив\'язана до фізичної структури; вона об\'єднує ресурси університетів, ' +
    'дослідницьких інститутів, компаній через грід-middleware. ' +
    'Учасники зберігають локальний контроль і можуть підключатись/відключатись динамічно.'
  ));
  
  content.push(heading2('47.3. Характеристики Grid'));
  content.push(bullet('Гетерогенність: різні ОС, апаратні платформи, ПЗ'));
  content.push(bullet('Географічна розподіленість: вузли по всьому світу'));
  content.push(bullet('Динамічність: ресурси входять і виходять з системи'));
  content.push(bullet('Масштабованість: від десятків до мільйонів вузлів (BOINC: ~800k активних хостів)'));
  content.push(bullet('Безпека: різні організаційні домени, потрібна PKI-автентифікація'));
  
  content.push(heading2('47.4. Застосування Grid'));
  content.push(bullet('CERN LHC Computing Grid: обробка даних Великого адронного колайдера'));
  content.push(bullet('BOINC (SETI@home, Folding@home): добровільні обчислення'));
  content.push(bullet('EGI (European Grid Infrastructure): наукові дослідження у Європі'));
  content.push(bullet('Фармацевтичні компанії: молекулярне моделювання ліків'));
  content.push(separator());
  
  // ══════════════════════════════════════════════
  // ПИТАННЯ 49
  // ══════════════════════════════════════════════
  content.push(heading1('Питання 49. Програмне забезпечення грід-технологій'));
  
  content.push(heading2('49.1. Шари ПЗ грід-системи'));
  content.push(para(
    'Архітектура ПЗ грід-системи складається з кількох шарів: ' +
    'Fabric layer (апаратні ресурси та локальні ОС), Connectivity layer (мережеві протоколи, безпека), ' +
    'Resource layer (управління одним ресурсом), Collective layer (управління набором ресурсів), ' +
    'Application layer (прикладні програми кінцевих користувачів).'
  ));
  
  content.push(heading2('49.2. Globus Toolkit'));
  content.push(para(
    'Globus Toolkit — де-факто стандарт middleware для Grid. Компоненти: ' +
    'GRAM (Grid Resource Allocation and Management) — подання і управління задачами; ' +
    'GridFTP — оптимізована передача файлів; ' +
    'GSI (Grid Security Infrastructure) — PKI-автентифікація на основі X.509 сертифікатів; ' +
    'MDS (Monitoring and Discovery Service) — пошук і моніторинг ресурсів.'
  ));
  
  content.push(heading2('49.3. BOINC (Berkeley Open Infrastructure for Network Computing)'));
  content.push(para(
    'BOINC — платформа для добровільних розподілених обчислень. Дозволяє науковим проектам ' +
    'використовувати обчислювальні потужності мільйонів добровольців. Клієнт BOINC ' +
    'запускається у фоні, виконуючи робочі пакети (work units), і повертає результати серверу. ' +
    'Забезпечує відмовостійкість: якщо результат не повернено, задача перевидається.'
  ));
  
  content.push(heading2('49.4. Cloud vs Grid'));
  content.push(para(
    'Хмарні обчислення (Cloud Computing) частково замінили Grid у комерційній сфері: ' +
    'AWS, Azure, GCP надають обчислювальні ресурси за моделлю pay-as-you-go, з простішим ' +
    'управлінням та кращою комерційною підтримкою. Grid залишається важливим для наукових ' +
    'досліджень, де потрібно об\'єднати ресурси кількох організацій без єдиного власника.'
  ));
  
  content.push(heading2('49.5. Технологічний стек для Grid задач у Java'));
  content.push(bullet('JavaGAT (Java Grid Application Toolkit): Java API для роботи з Grid ресурсами'));
  content.push(bullet('SAGA (Simple API for Grid Applications): стандартний API від OGF'));
  content.push(bullet('Hazelcast, Apache Ignite: In-memory data grids для Java'));
  content.push(bullet('Spring Batch + Cloud: для масштабованої пакетної обробки'));
  content.push(bullet('Apache Hadoop, Spark: MapReduce парадигма для обробки великих даних'));
  content.push(separator());
  
  // ─────────────────────────────────────────────
  // ПОБУДОВА ДОКУМЕНТУ
  // ─────────────────────────────────────────────
  
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 24 },
        },
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '2E4057' },
          paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '2E75B6' },
          paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '\u2022',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: { indent: { left: 720, hanging: 360 } },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1134, right: 850, bottom: 1134, left: 1134 },
          },
        },
        children: [
          // Титульна сторінка
          new Paragraph({
            children: [new TextRun({ text: '', size: 24 })],
            spacing: { before: 2000, after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Технологія паралельних обчислень',
                bold: true,
                size: 48,
                font: 'Arial',
                color: '2E4057',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 400 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Відповіді на екзаменаційні питання',
                size: 32,
                font: 'Arial',
                color: '666666',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Мова програмування: Java',
                size: 28,
                font: 'Arial',
                color: '888888',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 3000 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: '25 питань • Детальні відповіді з прикладами коду',
                size: 24,
                font: 'Arial',
                color: '999999',
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
          // Розрив сторінки перед першим питанням
          new Paragraph({
            children: [new PageBreak()],
          }),
          ...content,
        ],
      },
    ],
  });
  
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync('./docs/uv3.docx', buffer);
  });