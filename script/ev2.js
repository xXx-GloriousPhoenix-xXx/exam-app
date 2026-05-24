// Запуск: node generate_exam_docx.js
// Требует: npm install docx
// Результат: exam_parallel_computing.docx

const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    AlignmentType, LevelFormat, PageNumber, PageBreak,
    BorderStyle, WidthType
  } = require('docx');
  const fs = require('fs');
  
  // ─── Helpers ────────────────────────────────────────────────────────────────
  
  function h1(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
      children: [new TextRun({ text, bold: true, size: 32, font: 'Arial' })]
    });
  }
  
  function h2(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 300, after: 160 },
      children: [new TextRun({ text, bold: true, size: 28, font: 'Arial' })]
    });
  }
  
  function para(text, opts = {}) {
    return new Paragraph({
      spacing: { before: 80, after: 80 },
      alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
      children: [new TextRun({ text, size: 22, font: 'Arial', ...opts.run })]
    });
  }
  
  function bullet(text) {
    return new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, size: 22, font: 'Arial' })]
    });
  }
  
  function code(text) {
    return new Paragraph({
      spacing: { before: 60, after: 60 },
      indent: { left: 720 },
      children: [new TextRun({ text, size: 20, font: 'Courier New', color: '2E4057' })]
    });
  }
  
  function pageBreak() {
    return new Paragraph({ children: [new PageBreak()] });
  }
  
  function bold(text) {
    return new TextRun({ text, bold: true, size: 22, font: 'Arial' });
  }
  
  function mixed(...runs) {
    return new Paragraph({
      spacing: { before: 80, after: 80 },
      children: runs
    });
  }
  
  // ─── CONTENT ─────────────────────────────────────────────────────────────────
  
  const sections_content = [];
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 2
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 2. Класифікація паралельних обчислювальних систем'),
  
    h2('Що таке паралельна обчислювальна система?'),
    para('Паралельна обчислювальна система — це сукупність обчислювальних елементів (процесорів, ядер, вузлів), які здатні одночасно виконувати частини однієї задачі або різні задачі з метою прискорення обчислень або обробки великих обсягів даних.', { justify: true }),
    para('Основна мета паралелізму — скорочення часу розв\'язку задачі шляхом розподілу роботи між кількома виконавцями, що виконують обчислення одночасно.', { justify: true }),
  
    h2('Класифікація Флінна (Flynn\'s Taxonomy)'),
    para('Найпоширенішою класифікацією є таксономія Флінна (1966), яка розрізняє системи за двома потоками: потоком команд (Instruction Stream) та потоком даних (Data Stream).'),
    bullet('SISD (Single Instruction, Single Data) — один потік команд, один потік даних. Класичний фон-Нейманівський процесор. Паралелізм відсутній. Приклад: стара однопроцесорна ЕОМ.'),
    bullet('SIMD (Single Instruction, Multiple Data) — одна команда застосовується одночасно до множини даних. Використовується в GPU, векторних процесорах, SSE/AVX-розширеннях у сучасних CPU. Приклад: обробка пікселів зображення.'),
    bullet('MISD (Multiple Instruction, Single Data) — кілька команд обробляють один потік даних. На практиці майже не зустрічається; теоретична концепція. Можливий приклад — конвеєрна обробка сигналів.'),
    bullet('MIMD (Multiple Instruction, Multiple Data) — кілька процесорів виконують різні команди над різними даними незалежно. Це найпоширеніший клас сучасних паралельних систем: кластери, багатопроцесорні сервери, Grid-системи.'),
  
    h2('Класифікація за архітектурою пам\'яті'),
    bullet('Системи зі спільною пам\'яттю (Shared Memory) — всі процесори мають доступ до єдиного адресного простору. Зв\'язок між процесорами відбувається через читання/запис у спільну пам\'ять. Підклас: SMP (Symmetric Multiprocessing) та NUMA (Non-Uniform Memory Access).'),
    bullet('Системи з розподіленою пам\'яттю (Distributed Memory) — кожен процесор має власну локальну пам\'ять. Обмін між процесорами — через передачу повідомлень (MPI). Приклад: кластери з тисячами вузлів.'),
    bullet('Гібридні системи — поєднання обох підходів. Типовий приклад: HPC-кластер, де кожен вузол є багатоядерним (спільна пам\'ять всередині вузла), а вузли з\'єднані мережею (розподілена пам\'ять між вузлами).'),
  
    h2('Класифікація за топологією з\'єднань'),
    bullet('Шина (Bus) — всі процесори підключені до єдиної шини. Проста реалізація, але виникає вузьке місце (bottle-neck) при збільшенні кількості процесорів.'),
    bullet('Кільце (Ring) — вузли з\'єднані послідовно в кільце. Діаметр мережі — N/2.'),
    bullet('Повний граф (Fully Connected) — кожен вузол з\'єднаний з кожним. Максимальна швидкість обміну, але висока вартість з\'єднань.'),
    bullet('Гіперкуб (Hypercube) — N-вимірний куб. n вузлів, де n = 2^k. Ефективна балансова мережа для MPI.'),
    bullet('Решітка (Mesh/Torus) — двовимірна або тривимірна решітка. Використовується в суперкомп\'ютерах (наприклад, BlueGene).'),
    bullet('Дерево (Fat Tree) — ієрархічна структура. Широко використовується у сучасних HPC-інфраструктурах з Infiniband.'),
  
    h2('Класифікація за масштабом паралелізму'),
    bullet('Бітовий паралелізм — одночасна обробка кількох бітів (наприклад, 64-розрядна операція замість 8-розрядної).'),
    bullet('Паралелізм команд (Instruction-Level Parallelism, ILP) — конвеєрне виконання, суперскалярність.'),
    bullet('Паралелізм даних (Data-Level Parallelism) — SIMD, векторні операції.'),
    bullet('Паралелізм задач (Task-Level Parallelism) — незалежні задачі виконуються паралельно на різних процесорах.'),
  
    h2('Приклади реальних систем'),
    bullet('Багатоядерний процесор Intel Core i9 — MIMD зі спільною пам\'яттю (SMP).'),
    bullet('GPU NVIDIA (CUDA) — SIMD/SIMT архітектура, тисячі потоків.'),
    bullet('Кластер на базі MPI — MIMD з розподіленою пам\'яттю.'),
    bullet('Apache Spark / Hadoop — розподілені обчислення на рівні програмного забезпечення.'),
    bullet('Суперкомп\'ютери (TOP500) — переважно гібридні MIMD-системи.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 4
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 4. Системи з загальною та розподіленою пам\'яттю'),
  
    h2('Системи зі спільною (загальною) пам\'яттю'),
    para('У системах зі спільною пам\'яттю всі процесори (ядра) мають доступ до єдиного глобального адресного простору. Це означає, що будь-який процесор може читати або записувати будь-яку комірку пам\'яті без явної пересилки даних.', { justify: true }),
  
    h2('Підвиди архітектур зі спільною пам\'яттю'),
    bullet('UMA (Uniform Memory Access) — симетричний багатопроцесорний доступ. Кожен процесор має однаковий час доступу до будь-якої комірки пам\'яті. Підходить для 2–32 ядер. Приклад: SMP-сервер.'),
    bullet('NUMA (Non-Uniform Memory Access) — неоднорідний доступ. Пам\'ять фізично розподілена між вузлами, але логічно є спільною. Час доступу до «своєї» пам\'яті (local) менший, ніж до «чужої» (remote). Приклад: багатопроцесорні сервери AMD EPYC з кількома сокетами.'),
    bullet('cc-NUMA (Cache Coherent NUMA) — підтримується когерентність кешу між вузлами. Найпоширеніший варіант у сучасних багатопроцесорних системах.'),
  
    h2('Проблеми спільної пам\'яті'),
    bullet('Перегони (Race Conditions) — некоректний результат через несинхронізований доступ до спільних даних.'),
    bullet('Когерентність кешу — при зміні даних одним процесором інші можуть мати застарілу копію у кеші. Вирішується протоколами MESI, MOESI.'),
    bullet('Масштабованість — зі збільшенням кількості процесорів шина пам\'яті стає вузьким місцем.'),
  
    h2('Синхронізація в Java (спільна пам\'ять)'),
    para('У Java багатопотоковість базується на моделі спільної пам\'яті. Потоки одного процесу поділяють heap (купу), але кожен має власний стек.'),
    code('// Синхронізація через synchronized'),
    code('public synchronized void increment() {'),
    code('    this.counter++;'),
    code('}'),
    code(''),
    code('// Атомарна змінна — без блокування'),
    code('AtomicInteger counter = new AtomicInteger(0);'),
    code('counter.incrementAndGet();'),
  
    h2('Системи з розподіленою пам\'яттю'),
    para('У системах з розподіленою пам\'яттю кожен вузол (процесор або комп\'ютер) має власну локальну пам\'ять, недоступну безпосередньо іншим вузлам. Обмін даними здійснюється виключно через передачу повідомлень (Message Passing) по мережі.', { justify: true }),
  
    h2('Характеристики розподіленої пам\'яті'),
    bullet('Масштабованість — система легко розширюється до тисяч і мільйонів вузлів.'),
    bullet('Відсутність проблем когерентності — кожен процесор працює зі своїми даними.'),
    bullet('Необхідність явного обміну — програміст сам керує пересилкою даних (send/receive).'),
    bullet('Накладні витрати на комунікацію — мережеві затримки та пропускна здатність обмежують ефективність.'),
  
    h2('MPI — стандарт для розподіленої пам\'яті'),
    para('Message Passing Interface (MPI) — це стандарт бібліотеки обміну повідомленнями. Основні операції:'),
    bullet('MPI_Send / MPI_Recv — блокуючий обмін (процес чекає завершення операції).'),
    bullet('MPI_Isend / MPI_Irecv — неблокуючий обмін (процес продовжує роботу).'),
    bullet('MPI_Bcast — розсилка від одного до всіх.'),
    bullet('MPI_Scatter / MPI_Gather — розподіл / збір даних.'),
    bullet('MPI_Reduce / MPI_Allreduce — редукція (сума, максимум тощо).'),
  
    h2('Порівняння підходів'),
    bullet('Спільна пам\'ять: легше програмувати, важче масштабувати, проблеми синхронізації. Технології: Java Threads, OpenMP, pthreads.'),
    bullet('Розподілена пам\'ять: складніше програмувати, добре масштабується, без гонок даних. Технології: MPI, Spark, Hadoop.'),
    bullet('Гібридна модель: MPI між вузлами + OpenMP/Java Threads всередині вузла. Найефективніший підхід для HPC.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 6
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 6. Багатопоточна технологія Java'),
  
    h2('Що таке потік у Java?'),
    para('Потік (Thread) у Java — це найменша одиниця виконання, що планується операційною системою. Потоки одного процесу поділяють код, статичні дані та heap, але мають власні стек, лічильник команд та регістри процесора.', { justify: true }),
    para('Java надає вбудовану підтримку багатопоточності через пакет java.lang (Thread, Runnable) та java.util.concurrent (Executor, Future, Lock тощо).', { justify: true }),
  
    h2('Способи створення потоку'),
    para('Спосіб 1 — успадкування від Thread:'),
    code('class MyThread extends Thread {'),
    code('    @Override'),
    code('    public void run() {'),
    code('        System.out.println("Thread: " + getName());'),
    code('    }'),
    code('}'),
    code('MyThread t = new MyThread();'),
    code('t.start(); // Запускає новий потік'),
    para('Спосіб 2 — реалізація Runnable (перевага: клас може наслідувати інший клас):'),
    code('Runnable task = () -> System.out.println("Lambda thread");'),
    code('Thread t = new Thread(task);'),
    code('t.start();'),
    para('Спосіб 3 — Callable + Future (дозволяє повертати результат і обробляти виключення):'),
    code('Callable<Integer> callable = () -> 42;'),
    code('ExecutorService exec = Executors.newSingleThreadExecutor();'),
    code('Future<Integer> future = exec.submit(callable);'),
    code('int result = future.get(); // Блокує до завершення'),
    code('exec.shutdown();'),
  
    h2('Пріоритети потоків'),
    para('Java дозволяє встановлювати пріоритет потоку від 1 (MIN_PRIORITY) до 10 (MAX_PRIORITY), за замовчуванням — 5 (NORM_PRIORITY). Пріоритет впливає на планувальник ОС, але не гарантує порядок виконання.'),
    code('Thread t = new Thread(() -> {});'),
    code('t.setPriority(Thread.MAX_PRIORITY);'),
    code('t.start();'),
  
    h2('Стани потоку (Thread Lifecycle)'),
    bullet('NEW — потік створено, але start() не викликано.'),
    bullet('RUNNABLE — потік виконується або готовий до виконання.'),
    bullet('BLOCKED — потік чекає на монітор (synchronized блок).'),
    bullet('WAITING — потік чекає без тайм-ауту (Object.wait(), Thread.join()).'),
    bullet('TIMED_WAITING — чекає з тайм-аутом (Thread.sleep(ms), wait(ms)).'),
    bullet('TERMINATED — виконання завершено.'),
  
    h2('Зупинка, призупинка та відновлення потоку'),
    bullet('Thread.sleep(ms) — тимчасово призупиняє потік (TIMED_WAITING).'),
    bullet('Thread.yield() — поступається процесором іншим потокам.'),
    bullet('Thread.join() — поточний потік чекає завершення вказаного.'),
    bullet('interrupt() — встановлює прапор переривання. Потік перевіряє isInterrupted() або отримує InterruptedException.'),
    code('Thread t = new Thread(() -> {'),
    code('    while (!Thread.currentThread().isInterrupted()) {'),
    code('        // Робота'),
    code('    }'),
    code('});'),
    code('t.start();'),
    code('t.interrupt(); // Сигналізує про зупинку'),
  
    h2('Безпека потоків (Thread Safety)'),
    para('Клас або метод є потокобезпечним, якщо при одночасному виклику кількома потоками поводиться коректно. Засоби забезпечення потокобезпечності:'),
    bullet('synchronized — блокує монітор об\'єкту або класу.'),
    bullet('volatile — гарантує видимість змін між потоками (без атомарності).'),
    bullet('AtomicInteger, AtomicReference — lock-free операції через CAS (Compare-And-Swap).'),
    bullet('ReentrantLock, ReadWriteLock — явні блокування з гнучким управлінням.'),
    bullet('Immutable objects — незмінні об\'єкти автоматично потокобезпечні.'),
    bullet('ThreadLocal — кожен потік має власну копію змінної.'),
  
    h2('Бібліотека java.util.concurrent'),
    bullet('Executors — фабрика для створення пулів потоків (newFixedThreadPool, newCachedThreadPool, newScheduledThreadPool).'),
    bullet('ExecutorService — управління пулом, submit, shutdown, awaitTermination.'),
    bullet('Future, CompletableFuture — асинхронні обчислення з можливістю ланцюжка.'),
    bullet('BlockingQueue (ArrayBlockingQueue, LinkedBlockingQueue) — потокобезпечна черга для Producer-Consumer.'),
    bullet('CountDownLatch, CyclicBarrier, Semaphore — синхронізаційні примітиви.'),
    bullet('ForkJoinPool — для рекурсивного розпаралелення задач (divide-and-conquer).'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 8
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 8. Стани потоку та переходи між станами'),
  
    h2('Шість станів потоку в Java'),
    para('Перелік Thread.State в Java визначає шість можливих станів потоку. Планувальник ОС та JVM визначають, коли відбувається перехід між станами.', { justify: true }),
  
    h2('NEW'),
    para('Потік створено через new Thread(), але метод start() ще не викликано. JVM виділила об\'єкт Thread, але ще не запустила відповідний системний потік.'),
    code('Thread t = new Thread(() -> {}); // Стан: NEW'),
  
    h2('RUNNABLE'),
    para('Після виклику start() потік переходить у RUNNABLE. Він або вже виконується на процесорі, або стоїть у черзі планувальника, готовий до виконання. У Java цей стан об\'єднує і "Ready" і "Running" стану ОС.'),
    code('t.start(); // Стан: RUNNABLE'),
  
    h2('BLOCKED'),
    para('Потік намагається увійти в synchronized-блок або метод, монітор якого зайнятий іншим потоком. Потік очікує, поки інший не звільнить монітор.'),
    code('synchronized (lock) { // Якщо lock зайнятий — BLOCKED'),
    code('    // critical section'),
    code('}'),
  
    h2('WAITING'),
    para('Потік чекає без тайм-ауту. Вихід можливий лише за явним сигналом від іншого потоку. Методи, що переводять у WAITING:'),
    bullet('Object.wait() — звільняє монітор, чекає notify()/notifyAll().'),
    bullet('Thread.join() — чекає завершення іншого потоку.'),
    bullet('LockSupport.park() — блокує потік.'),
  
    h2('TIMED_WAITING'),
    para('Схоже на WAITING, але з тайм-аутом. Потік автоматично прокинеться після вказаного часу або при отриманні сигналу.'),
    bullet('Thread.sleep(long millis)'),
    bullet('Object.wait(long timeout)'),
    bullet('Thread.join(long millis)'),
    bullet('LockSupport.parkNanos(long nanos)'),
  
    h2('TERMINATED'),
    para('Потік завершив виконання методу run() (нормально або через виключення). Об\'єкт Thread ще існує, але потік більше не може бути запущений.'),
  
    h2('Схема переходів між станами'),
    bullet('NEW → RUNNABLE: виклик start()'),
    bullet('RUNNABLE → BLOCKED: спроба захопити зайнятий монітор'),
    bullet('BLOCKED → RUNNABLE: монітор звільнено'),
    bullet('RUNNABLE → WAITING: wait(), join() без тайм-ауту'),
    bullet('WAITING → RUNNABLE: notify(), notifyAll(), join завершився'),
    bullet('RUNNABLE → TIMED_WAITING: sleep(ms), wait(ms), join(ms)'),
    bullet('TIMED_WAITING → RUNNABLE: тайм-аут або сигнал'),
    bullet('RUNNABLE → TERMINATED: run() завершився'),
  
    h2('Метод interrupt() та InterruptedException'),
    para('Виклик interrupt() встановлює прапор переривання потоку. Якщо потік знаходиться в стані WAITING або TIMED_WAITING (wait, sleep, join), він негайно отримає InterruptedException і повернеться у RUNNABLE. Потік у стані RUNNABLE може самостійно перевіряти isInterrupted().'),
    code('Thread worker = new Thread(() -> {'),
    code('    try {'),
    code('        Thread.sleep(5000);'),
    code('    } catch (InterruptedException e) {'),
    code('        System.out.println("Interrupted!");'),
    code('        Thread.currentThread().interrupt(); // відновлюємо прапор'),
    code('    }'),
    code('});'),
    code('worker.start();'),
    code('worker.interrupt(); // примусове переривання'),
  
    h2('Демонічні потоки (Daemon Threads)'),
    para('Daemon-потоки — фонові потоки, що автоматично завершуються, коли всі non-daemon потоки закінчили роботу. Типовий приклад — GC (Garbage Collector). Встановлюється до start():'),
    code('Thread daemon = new Thread(() -> { /* фонова задача */ });'),
    code('daemon.setDaemon(true);'),
    code('daemon.start();'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 10
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 10. Алгоритми паралельного множення матриць'),
  
    h2('Постановка задачі'),
    para('Матричне множення C = A × B для матриць розміром N×N має послідовну складність O(N³). Паралелізація дозволяє суттєво скоротити час обчислення за рахунок розподілу роботи між кількома потоками або процесами.', { justify: true }),
  
    h2('Базовий паралельний алгоритм (рядковий розподіл)'),
    para('Найпростіший підхід: кожен потік обчислює один або кілька рядків результуючої матриці C. Рядок i матриці C:'),
    para('C[i][j] = Σ A[i][k] * B[k][j], k = 0..N-1'),
    code('int rows = N / numThreads;'),
    code('for (int t = 0; t < numThreads; t++) {'),
    code('    final int start = t * rows;'),
    code('    final int end = (t == numThreads-1) ? N : start + rows;'),
    code('    executor.submit(() -> {'),
    code('        for (int i = start; i < end; i++)'),
    code('            for (int j = 0; j < N; j++)'),
    code('                for (int k = 0; k < N; k++)'),
    code('                    C[i][j] += A[i][k] * B[k][j];'),
    code('    });'),
    code('}'),
  
    h2('Стрічковий (Striped) алгоритм'),
    para('У стрічковому алгоритмі матриці A і B розбиваються на горизонтальні стрічки. Процесор p обчислює стрічку C_p, маючи доступ до відповідної стрічки A_p та всієї матриці B (або передає стрічки B по кільцю).', { justify: true }),
    bullet('Розмір стрічки: N/P рядків на процесор.'),
    bullet('Перевага: простота реалізації, добре масштабується для систем зі спільною пам\'яттю.'),
    bullet('Недолік: при розподіленій пам\'яті кожен процес потребує всю матрицю B — велике споживання пам\'яті.'),
  
    h2('Алгоритм Кеннона (Cannon\'s Algorithm)'),
    para('Призначений для систем з розподіленою пам\'яттю на решітці процесорів P×P. Матриці A і B ініціально зсуваються (skew), після чого виконується P кроків локального перемноження та циклічного зсуву.'),
    bullet('Кожен процесор зберігає блок A[i][j] та B[i][j] розміром (N/P)×(N/P).'),
    bullet('Початковий зсув: A[i][j] зсувається на i позицій ліворуч, B[i][j] — на j позицій вгору.'),
    bullet('На кожному кроці: C[i][j] += A[i][j] * B[i][j], потім A зсувається ліворуч, B — вгору.'),
    bullet('Складність: O(N³/P) при P процесорах, комунікаційна складність O(N²/√P).'),
  
    h2('Алгоритм Фокса (Fox\'s Algorithm)'),
    para('Детально розглядається в питанні 12. Суть: матриця A транслюється по рядках блоків, після чого виконується локальне перемноження та зсув блоків B вгору.'),
  
    h2('Оцінка ефективності'),
    bullet('Прискорення S(P) = T(1) / T(P). Ідеально S(P) = P.'),
    bullet('Ефективність E(P) = S(P) / P. Прагнемо до 1.'),
    bullet('Закон Амдала обмежує прискорення: S ≤ 1 / (α + (1-α)/P), де α — частка послідовного коду.'),
    bullet('При N = 1024, P = 16 потоків можна отримати прискорення ~12–14x (через накладні витрати на синхронізацію).'),
  
    h2('Реалізація з ForkJoinPool у Java'),
    code('class MatMulTask extends RecursiveAction {'),
    code('    int startRow, endRow;'),
    code('    double[][] A, B, C;'),
    code('    static final int THRESHOLD = 64;'),
    code('    @Override'),
    code('    protected void compute() {'),
    code('        if (endRow - startRow <= THRESHOLD) {'),
    code('            // Пряме обчислення'),
    code('            for (int i = startRow; i < endRow; i++)'),
    code('                for (int j = 0; j < C[0].length; j++)'),
    code('                    for (int k = 0; k < B.length; k++)'),
    code('                        C[i][j] += A[i][k] * B[k][j];'),
    code('        } else {'),
    code('            int mid = (startRow + endRow) / 2;'),
    code('            invokeAll(new MatMulTask(startRow, mid, A, B, C),'),
    code('                      new MatMulTask(mid, endRow, A, B, C));'),
    code('        }'),
    code('    }'),
    code('}'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 12
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 12. Алгоритм Фокса паралельного множення матриць'),
  
    h2('Призначення та умови застосування'),
    para('Алгоритм Фокса (Fox\'s Algorithm, також відомий як BSP-алгоритм або алгоритм 2D-розпаралелення) — це блоковий алгоритм множення матриць, призначений для сіток процесорів P×P при розподіленій пам\'яті. Він мінімізує обсяг комунікацій порівняно з наївним розподілом.', { justify: true }),
  
    h2('Початкові умови'),
    bullet('Матриці A, B, C розміром N×N.'),
    bullet('P = q² процесорів, розташованих у сітці q×q.'),
    bullet('Кожен процесор (i, j) зберігає блоки A_ij, B_ij, C_ij розміром (N/q)×(N/q).'),
  
    h2('Кроки алгоритму Фокса'),
    para('Алгоритм виконується в q ітераціях (l = 0, 1, ..., q-1):'),
    bullet('Крок 1 (Broadcast): Процесор (i, (i+l) mod q) транслює свій блок A_i,((i+l) mod q) всім процесорам рядка i. Тобто активний блок матриці A розсилається по всьому рядку процесорів.'),
    bullet('Крок 2 (Multiply): Кожен процесор (i,j) виконує локальне множення: C_ij += A_broadcast * B_ij.'),
    bullet('Крок 3 (Shift): Блоки матриці B зсуваються на одну позицію вгору по стовпцям: B_ij → B_(i-1, j). Процесор (0, j) отримує блок від (q-1, j).'),
    bullet('Після q ітерацій кожен процесор (i,j) містить готовий результатний блок C_ij.'),
  
    h2('Псевдокод алгоритму Фокса'),
    code('// Ініціалізація: розподіл блоків між процесорами'),
    code('for l = 0 to q-1:'),
    code('    // 1. Визначити джерело трансляції в рядку i'),
    code('    source = (i + l) mod q'),
    code('    // 2. Транслювати A[i][source] по рядку процесорів i'),
    code('    MPI_Bcast(A_block, root=source, comm=row_comm)'),
    code('    // 3. Локальне множення'),
    code('    C_block += A_broadcast * B_block'),
    code('    // 4. Циклічний зсув B вгору'),
    code('    MPI_Sendrecv(B_block → rank_up, recv from rank_down)'),
  
    h2('Комунікаційна складність'),
    bullet('На кожному з q кроків: трансляція блоку A розміром (N/q)² — обсяг O(N²/q).'),
    bullet('Зсув блоку B — обсяг O(N²/q²).'),
    bullet('Загальний обсяг комунікацій: O(N²/q) = O(N²/√P).'),
    bullet('Порівняно з алгоритмом з повним обміном, де комунікаційна складність O(N²), алгоритм Фокса значно ефективніший.'),
  
    h2('Обчислювальна складність'),
    bullet('Кожен процесор виконує q множень блоків (N/q)×(N/q), тобто q * (N/q)³ = N³/q² = N³/P операцій.'),
    bullet('Теоретичне прискорення: S(P) = P.'),
    bullet('Ефективність: E(P) = 1 (ідеально збалансований алгоритм).'),
  
    h2('Реалізація Java (симуляція логіки Фокса для спільної пам\'яті)'),
    code('// Блочна версія для потоків зі спільною пам\'яттю'),
    code('int q = (int) Math.sqrt(numThreads);'),
    code('int blockSize = N / q;'),
    code(''),
    code('for (int l = 0; l < q; l++) {'),
    code('    final int step = l;'),
    code('    List<Future<?>> futures = new ArrayList<>();'),
    code('    for (int i = 0; i < q; i++) {'),
    code('        for (int j = 0; j < q; j++) {'),
    code('            final int bi = i, bj = j;'),
    code('            final int sourceCol = (i + step) % q;'),
    code('            futures.add(executor.submit(() -> {'),
    code('                // Broadcast: беремо блок A[bi][sourceCol]'),
    code('                multiplyBlocks(A, B, C, bi, sourceCol, bj, blockSize);'),
    code('            }));'),
    code('        }'),
    code('    }'),
    code('    // Чекаємо завершення кроку'),
    code('    for (Future<?> f : futures) f.get();'),
    code('    // Shift B вгору (симуляція)'),
    code('    shiftBUp(B, q, blockSize);'),
    code('}'),
  
    h2('Переваги та недоліки'),
    bullet('Перевага: оптимальна комунікаційна складність O(N²/√P).'),
    bullet('Перевага: добре підходить для кластерів з MPI.'),
    bullet('Недолік: вимагає кількість процесів у вигляді повного квадрата q².'),
    bullet('Недолік: складніша реалізація порівняно зі стрічковим алгоритмом.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 14
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 14. Призупинка та відновлення роботи потоку. Переривання.'),
  
    h2('Thread.sleep() — тимчасова призупинка'),
    para('Метод Thread.sleep(long millis) переводить поточний потік у стан TIMED_WAITING на вказану кількість мілісекунд. Процесорний час при цьому звільняється для інших потоків.'),
    code('try {'),
    code('    Thread.sleep(1000); // пауза 1 секунда'),
    code('} catch (InterruptedException e) {'),
    code('    Thread.currentThread().interrupt(); // відновлюємо прапор'),
    code('}'),
    para('Важливо: sleep не звільняє монітори (synchronized блоки), на відміну від wait().'),
  
    h2('Object.wait() та notify()'),
    para('Метод wait() переводить потік у стан WAITING та звільняє монітор об\'єкту. Потік прокидається при виклику notify() або notifyAll() іншим потоком, що утримує той самий монітор.'),
    code('synchronized (lock) {'),
    code('    while (!condition) {'),
    code('        lock.wait(); // звільняє монітор, чекає'),
    code('    }'),
    code('    // Умова виконана — продовжуємо'),
    code('}'),
    code('// В іншому потоці:'),
    code('synchronized (lock) {'),
    code('    condition = true;'),
    code('    lock.notifyAll(); // прокидає всіх'),
    code('}'),
    para('Зверніть увагу: умову перевіряють у циклі while (не if), щоб захиститися від spurious wakeup.'),
  
    h2('Thread.join() — очікування завершення'),
    para('Метод join() блокує поточний потік до завершення вказаного потоку. Корисно для синхронізації результатів.'),
    code('Thread worker = new Thread(() -> compute());'),
    code('worker.start();'),
    code('worker.join(); // Основний потік чекає завершення worker'),
    code('System.out.println("Worker done");'),
  
    h2('Thread.yield() — поступка процесором'),
    para('Метод yield() дає підказку планувальнику, що поточний потік готовий поступитися процесором іншим потокам з тим самим або вищим пріоритетом. Поведінка залежить від ОС та JVM.'),
    code('Thread.yield(); // Рекомендація, не гарантія'),
  
    h2('LockSupport.park() / unpark()'),
    para('LockSupport — низькорівневий механізм блокування потоків, що використовується всередині java.util.concurrent. park() блокує потік, unpark(thread) прокидає його.'),
    code('Thread t = new Thread(() -> {'),
    code('    LockSupport.park(); // блокує'),
    code('    System.out.println("Resumed");'),
    code('});'),
    code('t.start();'),
    code('Thread.sleep(100);'),
    code('LockSupport.unpark(t); // прокидає'),
  
    h2('Переривання потоку (interrupt)'),
    para('Java не має методу примусового зупинення потоку (stop() застарів як небезпечний). Замість цього використовується кооперативний механізм переривання:'),
    bullet('thread.interrupt() — встановлює прапор переривання потоку.'),
    bullet('thread.isInterrupted() — перевіряє прапор (не скидає його).'),
    bullet('Thread.interrupted() — перевіряє і скидає прапор поточного потоку.'),
    bullet('Якщо потік у wait/sleep/join — негайно кидається InterruptedException, прапор скидається.'),
    code('Thread worker = new Thread(() -> {'),
    code('    while (!Thread.currentThread().isInterrupted()) {'),
    code('        try {'),
    code('            processData();'),
    code('            Thread.sleep(100);'),
    code('        } catch (InterruptedException e) {'),
    code('            Thread.currentThread().interrupt(); // відновлюємо прапор'),
    code('            break; // завершуємо роботу'),
    code('        }'),
    code('    }'),
    code('});'),
    code('worker.start();'),
    code('// Пізніше зупиняємо:'),
    code('worker.interrupt();'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 16
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 16. Синхронізація в паралельних обчисленнях'),
  
    h2('Навіщо потрібна синхронізація?'),
    para('Синхронізація — це механізм координації дій кількох потоків або процесів при доступі до спільних ресурсів (змінних, файлів, мережевих з\'єднань). Без синхронізації виникають перегони (race conditions), що призводять до непередбачуваних результатів.', { justify: true }),
  
    h2('Проблема перегонів'),
    para('Приклад: два потоки інкрементують лічильник:'),
    code('// Потік 1: читає counter=5, потік 2: читає counter=5'),
    code('// Потік 1: пише counter=6, потік 2: пише counter=6'),
    code('// Результат: 6 замість 7 — втрачена операція!'),
    para('Інкремент counter++ складається з трьох операцій: read-modify-write. Між ними може перемикнутися інший потік.'),
  
    h2('Ключове слово synchronized'),
    para('synchronized гарантує, що лише один потік одночасно виконує захищений блок коду. Реалізується через монітор об\'єкту.'),
    code('// Synchronized метод — монітор this'),
    code('public synchronized void increment() { counter++; }'),
    code(''),
    code('// Synchronized блок — монітор довільного об\'єкту'),
    code('synchronized (lock) {'),
    code('    counter++;'),
    code('}'),
    para('Монітор — це механізм взаємного виключення (mutex), вбудований у кожен Java-об\'єкт. Потоки, що не отримали монітор, переходять у стан BLOCKED.'),
  
    h2('volatile — видимість змін'),
    para('Ключове слово volatile гарантує, що зміна змінної одним потоком одразу видима іншим (без кешування в регістрах). Не гарантує атомарності складених операцій (read-modify-write).'),
    code('volatile boolean running = true;'),
    code('// Потік 1: running = false; — одразу видно потоку 2'),
    code('// Потік 2: while (running) { ... }'),
  
    h2('Атомарні класи'),
    para('Пакет java.util.concurrent.atomic надає класи з атомарними операціями через CAS (Compare-And-Swap) — без блокувань.'),
    code('AtomicInteger counter = new AtomicInteger(0);'),
    code('counter.incrementAndGet(); // атомарно, без synchronized'),
    code('counter.compareAndSet(expected, newValue); // CAS операція'),
  
    h2('ReentrantLock — явні блокування'),
    para('ReentrantLock надає більше гнучкості порівняно з synchronized: спроба захоплення з тайм-аутом, переривана блокада, умовні змінні.'),
    code('ReentrantLock lock = new ReentrantLock();'),
    code('lock.lock();'),
    code('try {'),
    code('    // Критична секція'),
    code('} finally {'),
    code('    lock.unlock(); // ЗАВЖДИ у finally!'),
    code('}'),
  
    h2('Умовні змінні (Condition)'),
    code('Condition notFull = lock.newCondition();'),
    code('Condition notEmpty = lock.newCondition();'),
    code(''),
    code('// Producer:'),
    code('lock.lock();'),
    code('try {'),
    code('    while (queue.isFull()) notFull.await();'),
    code('    queue.add(item);'),
    code('    notEmpty.signal();'),
    code('} finally { lock.unlock(); }'),
  
    h2('Семафор (Semaphore)'),
    para('Семафор — лічильниковий примітив синхронізації. Дозволяє обмежити кількість потоків, що одночасно мають доступ до ресурсу.'),
    code('Semaphore sem = new Semaphore(3); // До 3 потоків одночасно'),
    code('sem.acquire(); // Декрементує лічильник'),
    code('try {'),
    code('    useResource();'),
    code('} finally {'),
    code('    sem.release(); // Інкрементує лічильник'),
    code('}'),
  
    h2('CountDownLatch та CyclicBarrier'),
    bullet('CountDownLatch — одноразовий лічильник. Один або кілька потоків чекають, поки countDown() не буде викликано N разів.'),
    bullet('CyclicBarrier — бар\'єр: всі потоки чекають одне одного, потім продовжують. Може використовуватися повторно (cyclic).'),
    code('CyclicBarrier barrier = new CyclicBarrier(4);'),
    code('// Кожен потік:'),
    code('barrier.await(); // Чекає всіх 4 потоків'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 18
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 18. Блокування об\'єкту'),
  
    h2('Що таке монітор об\'єкту?'),
    para('Кожен Java-об\'єкт містить вбудований монітор (intrinsic lock або monitor lock). Монітор — це інструмент взаємного виключення: лише один потік може утримувати монітор конкретного об\'єкту в будь-який момент часу.', { justify: true }),
  
    h2('Захоплення монітору через synchronized'),
    bullet('synchronized method: потік захоплює монітор this при вході і звільняє при виході.'),
    bullet('synchronized (obj): потік захоплює монітор obj.'),
    bullet('static synchronized method: монітор — об\'єкт Class (не екземпляр).'),
    code('class Counter {'),
    code('    private int value = 0;'),
    code('    private final Object lock = new Object();'),
    code(''),
    code('    public void increment() {'),
    code('        synchronized (lock) { // Захоплює монітор lock'),
    code('            value++;'),
    code('        } // Звільняє монітор автоматично'),
    code('    }'),
    code('}'),
  
    h2('Реентерабельність монітора'),
    para('Java-монітор реентерабельний (reentrant): якщо потік вже утримує монітор об\'єкту, він може повторно увійти в synchronized-блок цього ж об\'єкту без блокування. Це дозволяє synchronized-методам викликати інші synchronized-методи того ж об\'єкту.'),
    code('public synchronized void methodA() {'),
    code('    methodB(); // Можна — поток вже має монітор'),
    code('}'),
    code('public synchronized void methodB() { ... }'),
  
    h2('Блокування на рівні класу'),
    code('class MyClass {'),
    code('    public static synchronized void staticMethod() {'),
    code('        // Монітор: MyClass.class'),
    code('    }'),
    code('    public synchronized void instanceMethod() {'),
    code('        // Монітор: this'),
    code('    }'),
    code('}'),
    para('Монітор класу та монітор екземпляра — різні об\'єкти. Один не блокує інший.'),
  
    h2('ReadWriteLock — оптимізація для читань'),
    para('Якщо об\'єкт часто читається і рідко змінюється, ReadWriteLock значно збільшує продуктивність: кілька читачів можуть утримувати read lock одночасно, але write lock — виключний.'),
    code('ReadWriteLock rwLock = new ReentrantReadWriteLock();'),
    code(''),
    code('// Читання (конкурентне):'),
    code('rwLock.readLock().lock();'),
    code('try { return data; }'),
    code('finally { rwLock.readLock().unlock(); }'),
    code(''),
    code('// Запис (виключне):'),
    code('rwLock.writeLock().lock();'),
    code('try { data = newValue; }'),
    code('finally { rwLock.writeLock().unlock(); }'),
  
    h2('StampedLock — оптимістичне блокування'),
    para('StampedLock (Java 8+) додає оптимістичне читання: перевіряється, чи не було змін без захоплення блокування. Найвища продуктивність при переважно читальному навантаженні.'),
    code('StampedLock sl = new StampedLock();'),
    code('long stamp = sl.tryOptimisticRead();'),
    code('int x = data; // Читаємо без блокування'),
    code('if (!sl.validate(stamp)) { // Якщо були зміни'),
    code('    stamp = sl.readLock();  // Беремо звичайний read lock'),
    code('    try { x = data; }'),
    code('    finally { sl.unlockRead(stamp); }'),
    code('}'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 20
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 20. Дедлок та інші небажані стани паралельної програми'),
  
    h2('Дедлок (Deadlock)'),
    para('Дедлок — стан, коли два або більше потоків нескінченно очікують ресурсів, якими поперемінно утримують один одного. Система "заморожується" без можливості продовжити роботу.', { justify: true }),
    para('Класичний приклад:'),
    code('// Потік 1:                    // Потік 2:'),
    code('lock(A);                       lock(B);'),
    code('lock(B); // чекає B           lock(A); // чекає A'),
    para('Умови Коффмана (всі чотири необхідні для дедлоку):'),
    bullet('Взаємне виключення (Mutual Exclusion) — ресурс може утримувати лише один потік.'),
    bullet('Утримання та очікування (Hold and Wait) — потік утримує один ресурс і чекає іншого.'),
    bullet('Відсутність примусового звільнення (No Preemption) — ресурс не може бути примусово відібраний.'),
    bullet('Циклічне очікування (Circular Wait) — існує цикл потоків P1→P2→...→Pk→P1.'),
  
    h2('Запобігання дедлоку'),
    bullet('Впорядкування блокувань: завжди захоплювати блокування в одному і тому ж порядку. Якщо всі потоки беруть A до B — цикл неможливий.'),
    bullet('tryLock з тайм-аутом: якщо блокування не вдалося захопити за час, звільнити вже захоплені.'),
    bullet('Алгоритм банкіра: виділяти ресурс лише якщо система залишається в безпечному стані.'),
    code('if (lock1.tryLock(100, TimeUnit.MILLISECONDS)) {'),
    code('    if (lock2.tryLock(100, TimeUnit.MILLISECONDS)) {'),
    code('        try { /* робота */ }'),
    code('        finally { lock2.unlock(); }'),
    code('    }'),
    code('    lock1.unlock();'),
    code('}'),
  
    h2('Лайвлок (Livelock)'),
    para('Лайвлок — потоки не заблоковані, але постійно реагують на дії один одного, не просуваючись. Аналогія: два люди в коридорі постійно поступаються місцем один одному.'),
    para('Рішення: випадкова затримка перед повторною спробою, пріоритизація потоків.'),
  
    h2('Голодування (Starvation)'),
    para('Голодування — потік постійно не отримує доступу до ресурсу через те, що інші потоки з вищим пріоритетом постійно його перехоплюють.'),
    para('Рішення: ReentrantLock(true) — fair lock (FIFO порядок), зниження пріоритету "жадібних" потоків.'),
    code('ReentrantLock fairLock = new ReentrantLock(true); // FIFO'),
  
    h2('Гонки даних (Race Conditions)'),
    para('Результат програми залежить від порядку виконання потоків (scheduling). Виникають при незахищеному доступі до спільних змінних.'),
    para('Рішення: synchronized, volatile, AtomicXxx, незмінні об\'єкти (Immutable).'),
  
    h2('Виявлення дедлоків у Java'),
    code('// За допомогою ThreadMXBean:'),
    code('ThreadMXBean tmx = ManagementFactory.getThreadMXBean();'),
    code('long[] deadlocked = tmx.findDeadlockedThreads();'),
    code('if (deadlocked != null) {'),
    code('    ThreadInfo[] info = tmx.getThreadInfo(deadlocked, true, true);'),
    code('    for (ThreadInfo ti : info) System.out.println(ti);'),
    code('}'),
    para('Також: jstack <pid> в командному рядку для отримання дамп потоків з діагностикою дедлоків.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 22
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 22. Локери та управління потоками'),
  
    h2('Чому явні блокування кращі за synchronized?'),
    para('java.util.concurrent.locks.Lock і його реалізації надають більш гнучке управління порівняно зі synchronized:'),
    bullet('Спроба захоплення з тайм-аутом (tryLock).'),
    bullet('Переривна блокада (lockInterruptibly).'),
    bullet('Різні умовні змінні (Condition) для одного блокування.'),
    bullet('Fair vs Non-fair режими.'),
    bullet('Підтримка ReadWriteLock.'),
  
    h2('ReentrantLock'),
    code('ReentrantLock lock = new ReentrantLock();'),
    code(''),
    code('// Стандартне блокування'),
    code('lock.lock();'),
    code('try { /* критична секція */ }'),
    code('finally { lock.unlock(); }'),
    code(''),
    code('// Блокування з тайм-аутом'),
    code('if (lock.tryLock(500, TimeUnit.MILLISECONDS)) {'),
    code('    try { /* критична секція */ }'),
    code('    finally { lock.unlock(); }'),
    code('} else { /* не вдалося */ }'),
    code(''),
    code('// Переривна блокада'),
    code('lock.lockInterruptibly(); // Кидає InterruptedException'),
  
    h2('Condition — умовні змінні'),
    para('Condition замінює wait()/notify(), але прив\'язана до конкретного Lock. Можна мати кілька Condition для одного Lock.'),
    code('Lock lock = new ReentrantLock();'),
    code('Condition notFull  = lock.newCondition();'),
    code('Condition notEmpty = lock.newCondition();'),
    code(''),
    code('// Producer:'),
    code('lock.lock();'),
    code('try {'),
    code('    while (buffer.isFull()) notFull.await();'),
    code('    buffer.put(item);'),
    code('    notEmpty.signal();'),
    code('} finally { lock.unlock(); }'),
    code(''),
    code('// Consumer:'),
    code('lock.lock();'),
    code('try {'),
    code('    while (buffer.isEmpty()) notEmpty.await();'),
    code('    return buffer.take();'),
    code('    notFull.signal();'),
    code('} finally { lock.unlock(); }'),
  
    h2('ReentrantReadWriteLock'),
    para('Оптимізація для читальних навантажень: необмежена кількість concurrent readers, але write lock — виключний.'),
    code('ReentrantReadWriteLock rwl = new ReentrantReadWriteLock();'),
    code('Lock r = rwl.readLock(), w = rwl.writeLock();'),
    code(''),
    code('Map<String, Object> cache = new HashMap<>();'),
    code(''),
    code('Object get(String key) {'),
    code('    r.lock();'),
    code('    try { return cache.get(key); }'),
    code('    finally { r.unlock(); }'),
    code('}'),
    code(''),
    code('void put(String key, Object value) {'),
    code('    w.lock();'),
    code('    try { cache.put(key, value); }'),
    code('    finally { w.unlock(); }'),
    code('}'),
  
    h2('Паттерн Producer-Consumer з Lock'),
    para('Producer-Consumer (Виробник-Споживач) — класичний паттерн синхронізації. Виробники додають елементи до буфера, споживачі забирають. BlockingQueue значно спрощує реалізацію.'),
    code('BlockingQueue<Task> queue = new ArrayBlockingQueue<>(100);'),
    code(''),
    code('// Producer:'),
    code('executor.submit(() -> {'),
    code('    while (true) queue.put(generateTask()); // блокує при повному'),
    code('});'),
    code(''),
    code('// Consumer:'),
    code('executor.submit(() -> {'),
    code('    while (true) process(queue.take()); // блокує при порожньому'),
    code('});'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 24
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 24. Інтерфейс Executor та бібліотечні класи'),
  
    h2('Проблема прямого управління потоками'),
    para('Ручне створення потоків (new Thread()) має недоліки: накладні витрати на створення/знищення, відсутність обмеження кількості потоків, складність управління результатами та помилками.', { justify: true }),
  
    h2('Executor та ExecutorService'),
    para('Інтерфейс Executor (java.util.concurrent) визначає один метод execute(Runnable). ExecutorService розширює його методами управління: submit, shutdown, awaitTermination.'),
    code('ExecutorService exec = Executors.newFixedThreadPool(4);'),
    code('exec.execute(() -> System.out.println("Task 1"));'),
    code('Future<String> f = exec.submit(() -> "Result");'),
    code('String result = f.get(); // блокує'),
    code('exec.shutdown();'),
    code('exec.awaitTermination(60, TimeUnit.SECONDS);'),
  
    h2('Фабричні методи Executors'),
    bullet('newFixedThreadPool(n) — фіксований пул з n потоками. Задачі чекають у LinkedBlockingQueue.'),
    bullet('newCachedThreadPool() — пул з динамічним числом потоків. Незайняті потоки живуть 60 сек.'),
    bullet('newSingleThreadExecutor() — один потік, послідовне виконання задач.'),
    bullet('newScheduledThreadPool(n) — підтримує планування за часом: schedule, scheduleAtFixedRate.'),
    bullet('newWorkStealingPool() — ForkJoinPool з parallelism = CPU cores.'),
  
    h2('Future та CompletableFuture'),
    para('Future представляє результат асинхронної операції:'),
    code('Future<Integer> future = executor.submit(() -> heavyCalc());'),
    code('// Інша робота...'),
    code('int result = future.get(5, TimeUnit.SECONDS); // Тайм-аут'),
    code('boolean done = future.isDone();'),
    code('future.cancel(true); // Перервати'),
    para('CompletableFuture (Java 8+) — ланцюжок асинхронних операцій:'),
    code('CompletableFuture.supplyAsync(() -> fetchData())'),
    code('    .thenApply(data -> process(data))'),
    code('    .thenAccept(result -> save(result))'),
    code('    .exceptionally(ex -> { log(ex); return null; });'),
  
    h2('FixedThreadPool для СМО (Система масового обслуговування)'),
    para('Система масового обслуговування (СМО) моделює обробку запитів: клієнти надходять (Producer), сервери обробляють (Consumer). FixedThreadPool ідеально підходить: N потоків = N серверів.'),
    code('int servers = 4;'),
    code('ExecutorService pool = Executors.newFixedThreadPool(servers);'),
    code('BlockingQueue<Request> queue = new LinkedBlockingQueue<>(100);'),
    code(''),
    code('// Генератор запитів:'),
    code('ScheduledExecutorService gen = Executors.newScheduledThreadPool(1);'),
    code('gen.scheduleAtFixedRate(() -> {'),
    code('    try { queue.put(new Request()); }'),
    code('    catch (InterruptedException e) { Thread.currentThread().interrupt(); }'),
    code('}, 0, 500, TimeUnit.MILLISECONDS);'),
    code(''),
    code('// Сервери (обробники):'),
    code('for (int i = 0; i < servers; i++) {'),
    code('    pool.submit(() -> {'),
    code('        while (true) {'),
    code('            Request req = queue.take();'),
    code('            Thread.sleep(200 + random.nextInt(300)); // Час обслуговування'),
    code('        }'),
    code('    });'),
    code('}'),
  
    h2('ForkJoinPool — розпаралелення методом розбиття'),
    para('ForkJoinPool реалізує парадигму divide-and-conquer. Work-stealing: вільні потоки "крадуть" задачі з черг зайнятих потоків.'),
    code('class SearchTask extends RecursiveTask<List<String>> {'),
    code('    List<File> files; String pattern;'),
    code('    static final int THRESHOLD = 10;'),
    code(''),
    code('    @Override'),
    code('    protected List<String> compute() {'),
    code('        if (files.size() <= THRESHOLD) {'),
    code('            return searchSequential(files, pattern);'),
    code('        }'),
    code('        int mid = files.size() / 2;'),
    code('        SearchTask left = new SearchTask(files.subList(0, mid), pattern);'),
    code('        SearchTask right = new SearchTask(files.subList(mid, files.size()), pattern);'),
    code('        left.fork(); // асинхронно'),
    code('        List<String> rightResult = right.compute(); // синхронно'),
    code('        List<String> leftResult = left.join(); // чекаємо'),
    code('        leftResult.addAll(rightResult);'),
    code('        return leftResult;'),
    code('    }'),
    code('}'),
    code(''),
    code('ForkJoinPool pool = ForkJoinPool.commonPool();'),
    code('List<String> results = pool.invoke(new SearchTask(allFiles, "TODO"));'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 26
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 26. Методи моделювання паралельних обчислень'),
  
    h2('Навіщо моделювати паралельні програми?'),
    para('Паралельні програми складно аналізувати безпосередньо через недетермінізм, взаємодію між процесами та складні залежності. Моделювання дозволяє дослідити поведінку системи до її повної реалізації, виявити дедлоки, гонки та вузькі місця.', { justify: true }),
  
    h2('Граф залежностей задач (Task Dependency Graph)'),
    para('Орієнтований ациклічний граф (DAG), де вершини — задачі, ребра — залежності (задача B може розпочатись лише після завершення A). Використовується для:'),
    bullet('Знаходження критичного шляху (CP) — мінімально можливого часу виконання.'),
    bullet('Оцінки максимального паралелізму.'),
    bullet('Планування задач (scheduling).'),
    para('Прискорення: S ≤ T_total / T_cp, де T_cp — час критичного шляху.'),
  
    h2('Мережі Петрі (Petri Nets)'),
    para('Графічна та математична мова для специфікації паралельних, асинхронних, розподілених систем. Детально — у питанні 28.'),
  
    h2('Automata-based моделювання'),
    para('Скінченні автомати (FSM) та їх розширення (Communicating Sequential Processes, CSP) моделюють взаємодію процесів через синхронні або асинхронні канали. Використовується у верифікації протоколів (SPIN model checker).'),
  
    h2('Queueing Theory (Теорія черг)'),
    para('Математична теорія для аналізу систем масового обслуговування (СМО). Формула Літтла: L = λW (середня кількість у системі = інтенсивність × середній час перебування). Моделі M/M/1, M/M/c, M/G/1 для різних розподілів часу надходження та обслуговування.'),
    bullet('M/M/c: c серверів, пуасонівський потік, експоненційний час обслуговування.'),
    bullet('Дозволяє розрахувати: середній час очікування, ймовірність відмови, коефіцієнт завантаження.'),
  
    h2('Дискретно-подієве моделювання (DES)'),
    para('Система моделюється як послідовність подій у часі. Використовується для симуляції складних паралельних систем (кластери, мережі, СМО). Інструменти: SimPy (Python), GPSS, Arena.'),
  
    h2('Аналітичні моделі продуктивності'),
    bullet('Закон Амдала: обмежує прискорення за частки послідовного коду.'),
    bullet('Закон Густафсона: прискорення зростає зі збільшенням задачі.'),
    bullet('Roofline Model: аналіз арифметичної інтенсивності vs пропускна здатність пам\'яті.'),
  
    h2('Профілювання та інструменти вимірювань'),
    bullet('Java Flight Recorder (JFR) + JDK Mission Control — детальний профіль JVM.'),
    bullet('VisualVM — моніторинг потоків, пам\'яті, CPU в реальному часі.'),
    bullet('Intel VTune / Perf — аналіз на рівні CPU (cache misses, branch mispredictions).'),
    bullet('JMH (Java Microbenchmark Harness) — мікробенчмаркінг з автоматичним JIT warmup.'),
    code('@Benchmark'),
    code('@BenchmarkMode(Mode.Throughput)'),
    code('public void testMethod() { /* ... */ }'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 28
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 28. Моделювання паралельних програм мережею Петрі'),
  
    h2('Що таке мережа Петрі?'),
    para('Мережа Петрі — це математична модель паралельних дискретних систем, запропонована Карлом Адамом Петрі (1962). Складається з графічних і математичних компонентів, що дозволяють моделювати паралелізм, синхронізацію та взаємодію процесів.', { justify: true }),
  
    h2('Основні елементи'),
    bullet('Місця (Places, кола) — умови або стани системи. Містять маркери (токени).'),
    bullet('Переходи (Transitions, прямокутники) — події або дії.'),
    bullet('Дуги (Arcs) — з\'єднують місця і переходи (двочастинний граф).'),
    bullet('Маркери (Tokens, точки в місцях) — позначають поточний стан (розмітку).'),
  
    h2('Правило спрацьовування переходу'),
    para('Перехід T є ввімкненим (enabled), якщо кожне вхідне місце містить не менше ваги відповідної дуги маркерів. При спрацьовуванні: маркери видаляються з вхідних місць і додаються у вихідні.'),
  
    h2('Властивості мереж Петрі'),
    bullet('Досяжність (Reachability) — чи можна досягти певного стану з початкового.'),
    bullet('Обмеженість (Boundedness) — чи обмежена кількість маркерів у місцях. k-обмежена мережа: у кожному місці не більше k маркерів.'),
    bullet('Живість (Liveness) — кожен перехід може спрацювати знову. Відповідає відсутності дедлоків.'),
    bullet('Зворотність (Reversibility) — можливість повернутись до початкового стану.'),
  
    h2('Моделювання паралелізму'),
    para('Кілька переходів можуть бути одночасно увімкненими і спрацьовувати паралельно. Це безпосередньо відповідає паралельному виконанню потоків.'),
    para('Приклад: моделювання двох незалежних потоків обчислень через дві незалежні гілки мережі.'),
  
    h2('Моделювання Producer-Consumer'),
    para('Мережа Петрі для буфера ємністю N:'),
    bullet('Place "empty slots" (N маркерів початково) — кількість вільних місць.'),
    bullet('Place "full slots" (0 маркерів початково) — кількість готових елементів.'),
    bullet('Transition "produce": споживає 1 з "empty slots", додає 1 до "full slots".'),
    bullet('Transition "consume": споживає 1 з "full slots", додає 1 до "empty slots".'),
    para('Такий граф коректно моделює блокування при повному буфері (empty=0) та при порожньому (full=0).'),
  
    h2('Моделювання мьютекса (Mutex)'),
    para('Місце "mutex" з одним маркером. Перехід "enter_cs" споживає маркер, "exit_cs" повертає. Лише один потік може мати маркер — взаємне виключення гарантоване.'),
  
    h2('Розширення мереж Петрі'),
    bullet('Кольорові мережі Петрі (CPN) — маркери мають типи (кольори), що дозволяє компактно моделювати складні системи з даними.'),
    bullet('Часові мережі Петрі (TPN) — переходам присвоюється час спрацьовування для аналізу продуктивності.'),
    bullet('Стохастичні мережі Петрі (SPN) — час спрацьовування задано розподілом ймовірностей (для аналізу СМО).'),
  
    h2('Застосування в паралельному програмуванні'),
    bullet('Верифікація протоколів взаємодії MPI-процесів.'),
    bullet('Аналіз відсутності дедлоків у системах з блокуваннями.'),
    bullet('Моделювання алгоритмів синхронізації (бар\'єри, семафори).'),
    bullet('Проектування обробника задач для ForkJoinPool.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 30
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 30. Алгоритми паралельного сумування та оцінка їх ефективності'),
  
    h2('Постановка задачі'),
    para('Дано масив з N елементів. Потрібно обчислити їхню суму паралельно. Послідовний алгоритм: O(N) операцій. Паралельний може зменшити час до O(log N) при N процесорах.', { justify: true }),
  
    h2('Алгоритм 1: Наївний розподіл'),
    para('Масив ділиться на P частин. Кожен потік обчислює локальну суму. Потім локальні суми підсумовуються послідовно.'),
    code('ExecutorService exec = Executors.newFixedThreadPool(P);'),
    code('int chunk = N / P;'),
    code('AtomicLong total = new AtomicLong(0);'),
    code('List<Future<?>> futures = new ArrayList<>();'),
    code('for (int t = 0; t < P; t++) {'),
    code('    int start = t * chunk, end = (t == P-1) ? N : start + chunk;'),
    code('    futures.add(exec.submit(() -> {'),
    code('        long local = 0;'),
    code('        for (int i = start; i < end; i++) local += arr[i];'),
    code('        total.addAndGet(local);'),
    code('    }));'),
    code('}'),
    code('for (Future<?> f : futures) f.get();'),
    para('Складність: T_p = O(N/P + P) — підсумовування P результатів послідовне.'),
  
    h2('Алгоритм 2: Деревовидне підсумовування (Reduction Tree)'),
    para('Ефективніший підхід: на кожному рівні пари потоків об\'єднують результати. Дерево з log₂(P) рівнів.'),
    bullet('Крок 1: P потоків, кожен рахує локальну суму (N/P елементів).'),
    bullet('Крок 2: P/2 пар потоків об\'єднують суми.'),
    bullet('Крок 3: P/4 пар...'),
    bullet('Крок log₂(P): 1 фінальна сума.'),
    para('Час: T_p = O(N/P) + O(log P). При P = N: T_p = O(log N).'),
    code('// ForkJoinPool автоматично реалізує дерево:'),
    code('class SumTask extends RecursiveTask<Long> {'),
    code('    long[] arr; int lo, hi;'),
    code('    protected Long compute() {'),
    code('        if (hi - lo <= THRESHOLD) {'),
    code('            long sum = 0;'),
    code('            for (int i = lo; i < hi; i++) sum += arr[i];'),
    code('            return sum;'),
    code('        }'),
    code('        int mid = (lo + hi) / 2;'),
    code('        SumTask left = new SumTask(arr, lo, mid);'),
    code('        SumTask right = new SumTask(arr, mid, hi);'),
    code('        left.fork();'),
    code('        return right.compute() + left.join();'),
    code('    }'),
    code('}'),
  
    h2('Алгоритм 3: Паралельне префіксне сумування (Scan)'),
    para('Prefix Sum: для кожного i обчислити суму arr[0..i]. Корисний для паралельних сортувань, компресії.'),
    bullet('Фаза "Up-sweep" (reduce): обчислюються часткові суми знизу вгору по дереву.'),
    bullet('Фаза "Down-sweep": розповсюдження сум зверху вниз. Результат — масив prefix sums.'),
    para('Складність: O(N) роботи, O(log N) часу при N/2 процесорах.'),
  
    h2('Паралельне сумування в Java Streams'),
    code('long sum = Arrays.stream(arr).parallel().sum();'),
    code('// або:'),
    code('long sum = LongStream.of(arr).parallel().reduce(0L, Long::sum);'),
    para('Java parallel streams використовують ForkJoinPool.commonPool() і автоматично розбивають дані.'),
  
    h2('Оцінка ефективності'),
    bullet('S(P) = T(1)/T(P) — прискорення (speedup).'),
    bullet('E(P) = S(P)/P — ефективність.'),
    bullet('При N=10^7, P=8: наївний алгоритм дає S≈6-7, деревовидний — S≈7-7.5 (обмеження Амдала через фазу збирання).'),
    bullet('Масштабованість (Scalability): для фіксованого N ефективність падає зі збільшенням P.'),
    bullet('Слабка масштабованість: якщо зі збільшенням P збільшувати N пропорційно, ефективність залишається постійною.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 32
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 32. Експериментальне дослідження ефективності паралельних обчислень'),
  
    h2('Основні метрики'),
    bullet('Час виконання T(P) — час від початку до кінця при P процесорах.'),
    bullet('Прискорення S(P) = T(1) / T(P) — у скільки разів швидше при P процесорах.'),
    bullet('Ефективність E(P) = S(P) / P — ідеально = 1 (100% завантаження процесорів).'),
    bullet('Вартість C(P) = T(P) * P — сумарний процесорний час. Оптимально C(P) ≈ T(1).'),
    bullet('Надлишок (Overhead) Q(P) = C(P) - T(1) — накладні витрати синхронізації та комунікації.'),
  
    h2('Закон Амдала'),
    para('Якщо частка f програми є послідовною (не може бути паралелізована), то максимальне прискорення при P процесорах:'),
    para('S(P) ≤ 1 / (f + (1-f)/P)'),
    para('При f = 0.05 (5% послідовного коду) і P → ∞: S_max = 1/0.05 = 20. Навіть нескінченно багато процесорів не дасть більше ніж 20-кратне прискорення.'),
  
    h2('Закон Густафсона'),
    para('Якщо зі збільшенням P збільшується і розмір задачі (слабка масштабованість), то ефективне прискорення:'),
    para('S(P) = P - f*(P-1)'),
    para('Оптимістичніший закон: великі задачі масштабуються краще.'),
  
    h2('Методологія проведення вимірювань'),
    bullet('Розігрів JVM (JIT Warmup): перші N ітерацій не вимірювати, щоб JIT скомпілював код.'),
    bullet('Кілька вимірювань: виконати benchmark ≥10 разів, обчислити середнє та стандартне відхилення.'),
    bullet('Ізоляція: вимикати фонові процеси, GC-паузи враховувати.'),
    bullet('Правильні одиниці: мікросекунди для коротких задач, секунди — для тривалих.'),
    bullet('System.nanoTime() — вимірювання часу в Java (не currentTimeMillis — нижча точність).'),
  
    h2('JMH — Java Microbenchmark Harness'),
    code('@BenchmarkMode(Mode.AverageTime)'),
    code('@OutputTimeUnit(TimeUnit.MILLISECONDS)'),
    code('@Warmup(iterations = 5)'),
    code('@Measurement(iterations = 10)'),
    code('@Fork(2)'),
    code('public class MatMulBenchmark {'),
    code('    @Benchmark'),
    code('    public void sequential() { seqMatMul(A, B, C, N); }'),
    code(''),
    code('    @Benchmark'),
    code('    public void parallel4() { parMatMul(A, B, C, N, 4); }'),
    code('}'),
  
    h2('Графіки ефективності'),
    para('Типові графіки для аналізу:'),
    bullet('S(P) від P: порівняти з ідеальним (лінійним) прискоренням.'),
    bullet('E(P) від P: показує, як ефективно використовуються ресурси.'),
    bullet('T(P) від N при фіксованому P: підтверджує теоретичну складність.'),
    bullet('Аномальне прискорення: S > P буває при кращому кешуванні у дрібних блоках.'),
  
    h2('Джерела накладних витрат'),
    bullet('Створення та знищення потоків — дорого. Рішення: пул потоків.'),
    bullet('Синхронізація (lock contention) — блокування гальмує продуктивність.'),
    bullet('False sharing — різні потоки змінюють різні змінні в одній cache line. Рішення: padding.'),
    bullet('Комунікаційні затримки (latency) в MPI.'),
    bullet('Незбалансованість навантаження (load imbalance) — частина потоків простоює.'),
    code('@Contended // Java 8+ — анотація для запобігання false sharing'),
    code('volatile long counter;'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 34
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 34. Проектування паралельних програм'),
  
    h2('Методологія PCAM'),
    para('Методологія PCAM (Ian Foster) — систематичний підхід до проектування паралельних алгоритмів, що складається з чотирьох кроків:'),
    bullet('P — Partitioning (Розбиття): декомпозиція задачі на дрібні підзадачі. Два види: за даними (data decomposition) та за функціями (functional decomposition).'),
    bullet('C — Communication (Комунікація): визначення потоків даних між підзадачами. Локальна vs глобальна комунікація, синхронна vs асинхронна.'),
    bullet('A — Agglomeration (Агломерація): об\'єднання дрібних підзадач у крупніші для зменшення накладних витрат на комунікацію.'),
    bullet('M — Mapping (Відображення): призначення агломерованих задач конкретним процесорам/потокам. Мета: мінімізувати комунікації та балансувати навантаження.'),
  
    h2('Декомпозиція за даними'),
    para('Одні й ті самі операції виконуються над різними частинами даних (SPMD — Single Program, Multiple Data):'),
    bullet('Рядкова декомпозиція матриці: кожен потік обробляє свій рядок/блок рядків.'),
    bullet('Стовпцева декомпозиція: кожен потік обробляє свій стовпець/блок стовпців.'),
    bullet('Блокова декомпозиція: двовимірний розподіл блоків (алгоритм Фокса, Кеннона).'),
  
    h2('Декомпозиція за задачами (Task Decomposition)'),
    para('Різні операції виконуються паралельно:'),
    bullet('Pipeline (Конвеєр): стадія 1 → стадія 2 → стадія 3. Кожна стадія обробляє свій елемент одночасно.'),
    bullet('Master-Worker: master розподіляє задачі пулу робочих потоків. Динамічна балансировка.'),
    bullet('Divide-and-Conquer: рекурсивний поділ (ForkJoinPool).'),
  
    h2('Шаблони паралельних обчислень'),
    bullet('Map: застосувати функцію до кожного елементу незалежно. Ідеально паралелізується. Java: parallelStream().map().'),
    bullet('Reduce: згортка елементів (сума, максимум). Деревовидний алгоритм.'),
    bullet('Scan (Prefix Sum): кожному елементу — результат функції над попередніми.'),
    bullet('Stencil: обчислення залежать від сусідніх елементів (теплопровідність, обробка зображень).'),
    bullet('Pipeline: послідовність стадій обробки. Потоки даних течуть від стадії до стадії.'),
  
    h2('Балансування навантаження (Load Balancing)'),
    bullet('Статичне: навантаження розподілено заздалегідь. Ефективно, якщо час виконання підзадач однаковий.'),
    bullet('Динамічне: потоки отримують нові задачі з черги після завершення поточних. ForkJoinPool використовує work-stealing.'),
    bullet('Work-Stealing: вільний потік "краде" задачу з кінця черги зайнятого потоку — мінімізує синхронізацію.'),
  
    h2('Принципи проектування ефективних паралельних програм'),
    bullet('Мінімізувати синхронізацію: кожен lock — джерело contention.'),
    bullet('Локальність даних: дані потоку повинні бути в "своєму" кеші. Уникати false sharing.'),
    bullet('Обрати правильну гранулярність: замалі задачі — великі накладні витрати; завеликі — поганий баланс.'),
    bullet('Immutable objects: незмінні об\'єкти не потребують синхронізації.'),
    bullet('ThreadLocal: унікальні копії даних для кожного потоку без синхронізації.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 36
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 36. Моделі пам\'яті паралельних обчислень'),
  
    h2('Модель пам\'яті Java (Java Memory Model, JMM)'),
    para('JMM визначає, як потоки взаємодіють через пам\'ять, які значення може бачити кожен потік, та коли зміни стають видимими. Специфіковано в JSR-133 (Java 5+).', { justify: true }),
  
    h2('Happens-Before (HB)'),
    para('Ключове поняття JMM. Якщо операція A happens-before операції B, то:'),
    bullet('A виконується до B.'),
    bullet('Результати A видимі для B.'),
    para('Правила Happens-Before:'),
    bullet('Program Order: у межах одного потоку кожна операція HB наступній.'),
    bullet('Monitor Lock: unlock(m) HB перед lock(m) (хронологічно пізнім).'),
    bullet('Volatile Write/Read: запис у volatile поле HB читання цього поля.'),
    bullet('Thread Start: Thread.start() HB усіх операцій нового потоку.'),
    bullet('Thread Join: завершення потоку HB Thread.join() у чекаючому потоці.'),
  
    h2('Видимість та volatile'),
    para('Без volatile/synchronized процесор може зберігати змінні у регістрах або кеші, не скидаючи у RAM. Інші потоки бачать застарілі значення.'),
    code('// Небезпечний прапор зупинки:'),
    code('boolean stop = false; // потік 2 може ніколи не побачити true'),
    code(''),
    code('// Правильно:'),
    code('volatile boolean stop = false;'),
  
    h2('Атомарність'),
    para('Гарантія: операції int, float та менші — атомарні. Але long та double (64-bit) — не гарантовано атомарні на 32-bit JVM. Рекомендується volatile long.'),
    code('volatile long counter; // безпечно'),
    code('// або:'),
    code('AtomicLong counter = new AtomicLong(0);'),
  
    h2('Переупорядкування (Reordering)'),
    para('JVM та CPU можуть переупорядковувати операції для оптимізації, якщо це не порушує HB. Це може дивувати у багатопоточному контексті. synchronized та volatile запобігають проблемному переупорядкуванню.'),
  
    h2('Safe Publication (Безпечна публікація об\'єктів)'),
    para('Об\'єкт правильно опублікований, якщо всі потоки бачать його поля коректно ініціалізованими. Способи безпечної публікації:'),
    bullet('Статичний ініціалізатор: ClassLoader гарантує безпечну публікацію.'),
    bullet('volatile поле: прямий запис у volatile.'),
    bullet('final поля: після конструктора final поля гарантовано видимі.'),
    bullet('synchronized блок: запис та читання в одному моніторі.'),
    bullet('Concurrent collections (ConcurrentHashMap, CopyOnWriteArrayList).'),
  
    h2('Модель пам\'яті C++11/OpenMP'),
    para('Сучасний C++ та OpenMP мають схожу концепцію HB через std::atomic та memory_order. OpenMP надає директиви #pragma omp flush для синхронізації кешу між потоками.'),
  
    h2('NUMA та локальність пам\'яті'),
    para('У NUMA-системах час доступу до пам\'яті залежить від фізичного вузла. Для максимальної продуктивності потік має працювати з пам\'яттю того ж вузла (thread affinity + memory binding через numactl).'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 38
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 38. Стандарт Message Passing Interface (MPI)'),
  
    h2('Що таке MPI?'),
    para('MPI (Message Passing Interface) — стандарт бібліотеки обміну повідомленнями для паралельних обчислень на системах з розподіленою пам\'яттю. Визначає API, незалежне від мови програмування. Основні реалізації: OpenMPI, MPICH, Intel MPI.', { justify: true }),
    para('Концепція: P рівноправних процесів, кожен має власну пам\'ять та унікальний rank (0 до P-1). Взаємодія — виключно через явну передачу повідомлень.'),
  
    h2('Базові операції MPI'),
    bullet('MPI_Init / MPI_Finalize — ініціалізація та завершення.'),
    bullet('MPI_Comm_rank — отримати свій ранг.'),
    bullet('MPI_Comm_size — отримати кількість процесів.'),
    bullet('MPI_Send — блокуючий відправник.'),
    bullet('MPI_Recv — блокуючий отримувач.'),
    bullet('MPI_Isend / MPI_Irecv — неблокуючий відправник/отримувач.'),
    bullet('MPI_Wait / MPI_Waitall — очікування завершення неблокуючих операцій.'),
  
    h2('Реалізація MPI в Java'),
    para('Бібліотеки: mpi4j, MPJ Express (MPJ). Приклад з MPJ Express:'),
    code('import mpi.*;'),
    code(''),
    code('public class MPIHello {'),
    code('    public static void main(String[] args) throws Exception {'),
    code('        MPI.Init(args);'),
    code('        int rank = MPI.COMM_WORLD.Rank();'),
    code('        int size = MPI.COMM_WORLD.Size();'),
    code('        System.out.println("Process " + rank + " of " + size);'),
    code('        MPI.Finalize();'),
    code('    }'),
    code('}'),
  
    h2('Блокуючі методи P2P'),
    para('MPI_Send блокує, доки повідомлення не буде безпечно буферизовано або отримано. MPI_Recv блокує, доки повідомлення не прийде.'),
    code('// Процес 0 відправляє, процес 1 отримує:'),
    code('if (rank == 0) {'),
    code('    int[] data = {1, 2, 3, 4};'),
    code('    MPI.COMM_WORLD.Send(data, 0, 4, MPI.INT, 1, 0);'),
    code('} else if (rank == 1) {'),
    code('    int[] buf = new int[4];'),
    code('    MPI.COMM_WORLD.Recv(buf, 0, 4, MPI.INT, 0, 0);'),
    code('}'),
    para('Небезпека: взаємне блокування якщо обидва процеси спочатку Send, потім Recv.'),
  
    h2('Неблокуючі методи'),
    code('Request req = MPI.COMM_WORLD.Isend(buf, 0, N, MPI.DOUBLE, dest, tag);'),
    code('// Паралельна робота...'),
    code('req.Wait(); // Дочекатися завершення'),
  
    h2('Типи даних MPI'),
    bullet('MPI.INT, MPI.DOUBLE, MPI.FLOAT, MPI.CHAR, MPI.BYTE — примітивні.'),
    bullet('MPI_Type_contiguous, MPI_Type_vector — похідні типи для нерегулярних даних.'),
    bullet('MPI_Pack / MPI_Unpack — ручне упаковування різнорідних даних.'),
  
    h2('Топологія процесів'),
    bullet('MPI_Cart_create — декартова решітка процесів (для матричних алгоритмів).'),
    bullet('MPI_Cart_shift — знаходження сусідніх рангів у решітці.'),
    bullet('MPI_Graph_create — довільний граф зв\'язків.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 40
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 40. Колективні методи MPI та їх застосування'),
  
    h2('Що таке колективні операції?'),
    para('Колективні операції MPI — це комунікаційні операції, в яких беруть участь усі процеси комунікатора. На відміну від P2P (point-to-point), кожен процес повинен викликати колективну операцію. Оптимізовані реалізації значно ефективніші за еквівалент на P2P.', { justify: true }),
  
    h2('MPI_Bcast — трансляція'),
    para('Один root-процес надсилає дані всім іншим:'),
    code('int[] data = (rank == 0) ? new int[]{42} : new int[1];'),
    code('MPI.COMM_WORLD.Bcast(data, 0, 1, MPI.INT, 0);'),
    para('Застосування: розсилка матриці A у стрічковому алгоритмі множення; розсилка параметрів задачі від root.'),
  
    h2('MPI_Scatter та MPI_Gather'),
    para('Scatter: розподіл масиву від root рівними частинами між усіма процесами. Gather: зворотня операція — збір.'),
    code('// Root має масив N елементів, роздає кожному N/P:'),
    code('MPI.COMM_WORLD.Scatter(sendBuf, 0, N/P, MPI.INT, 0,'),
    code('                       recvBuf, 0, N/P, MPI.INT, 0);'),
    code('// Збір результатів:'),
    code('MPI.COMM_WORLD.Gather(recvBuf, 0, N/P, MPI.INT,'),
    code('                      result, 0, N/P, MPI.INT, 0);'),
  
    h2('MPI_Allgather та MPI_Alltoall'),
    bullet('Allgather: кожен процес збирає дані від всіх (результат однаковий у всіх).'),
    bullet('Alltoall: кожен процес надсилає різні частини своїх даних кожному іншому процесу. Використовується при транспонуванні матриць.'),
  
    h2('MPI_Reduce та MPI_Allreduce'),
    para('Reduce: всі процеси мають локальні дані, операція виконується і результат надходить до root.'),
    code('double localSum = computeLocalSum();'),
    code('double[] totalSum = new double[1];'),
    code('MPI.COMM_WORLD.Reduce(new double[]{localSum}, 0, totalSum, 0,'),
    code('                      1, MPI.DOUBLE, MPI.SUM, 0);'),
    para('Allreduce: як Reduce, але результат отримують всі процеси. Ключова операція у паралельних алгоритмах де всі потребують глобального результату (наприклад, глобального максимуму для нормалізації).'),
    code('MPI.COMM_WORLD.Allreduce(local, total, 1, MPI.DOUBLE, MPI.SUM);'),
  
    h2('Операції редукції MPI'),
    bullet('MPI.SUM — сума.'),
    bullet('MPI.MAX / MPI.MIN — максимум / мінімум.'),
    bullet('MPI.PROD — добуток.'),
    bullet('MPI.LAND / MPI.LOR — логічне AND / OR.'),
    bullet('MPI.MAXLOC / MPI.MINLOC — значення і ранг процесу з максимумом/мінімумом.'),
  
    h2('MPI_Barrier — синхронізаційний бар\'єр'),
    para('Всі процеси зупиняються до тих пір, поки останній не досягне бар\'єру. Використовується для синхронізації фаз алгоритму та коректного вимірювання часу.'),
    code('MPI.COMM_WORLD.Barrier();'),
    code('long startTime = System.nanoTime();'),
    code('// ... обчислення ...'),
    code('MPI.COMM_WORLD.Barrier();'),
    code('long elapsed = System.nanoTime() - startTime;'),
  
    h2('Застосування у матричному множенні'),
    para('Стрічковий алгоритм:'),
    bullet('Bcast матриці B (або її частини) від root до всіх.'),
    bullet('Scatter рядків матриці A між процесами.'),
    bullet('Кожен процес обчислює локальну частину C.'),
    bullet('Gather результатів до root.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 42
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 42. Розробка ефективних паралельних алгоритмів в OpenMPI'),
  
    h2('Принципи ефективності в MPI'),
    para('Ефективність MPI-програми визначається балансом між часом обчислень та часом комунікацій. Мета — максимально перекрити (overlap) комунікації з обчисленнями, мінімізувати кількість та обсяг повідомлень.', { justify: true }),
  
    h2('Перекриття комунікацій та обчислень'),
    para('Неблокуючі операції дозволяють процесору обчислювати, поки дані передаються мережею:'),
    code('// Неефективно (блокуючі):'),
    code('sendData();        // Чекаємо відправки'),
    code('compute(localData); // Тільки потім обчислюємо'),
    code(''),
    code('// Ефективно (неблокуючі):'),
    code('Request req = Isend(data);  // Відправляємо асинхронно'),
    code('compute(localData);          // Одночасно обчислюємо'),
    code('req.Wait();                  // Дочекатись лише наприкінці'),
  
    h2('Мінімізація комунікацій'),
    bullet('Агрегація повідомлень: краще одне велике повідомлення, ніж багато дрібних (latency > bandwidth cost).'),
    bullet('Локальний обмін: використовувати топологію (декартова решітка) для комунікацій лише між сусідами.'),
    bullet('Редукція кількості Bcast: кешувати отримані дані, не запитувати знову.'),
    bullet('Derived datatypes: передавати нерегулярні дані (рядки, діагоналі матриці) без копіювання.'),
  
    h2('Балансування навантаження'),
    para('Нерівномірне навантаження веде до простою швидших процесів. Статичний розподіл ефективний для рівномірних задач. Для нерівномірних — динамічний розподіл через Master-Worker:'),
    bullet('Master (rank 0) зберігає чергу задач.'),
    bullet('Worker запитує нову задачу після завершення поточної.'),
    bullet('Master надсилає задачу або сигнал "завершено".'),
  
    h2('Вибір алгоритмів редукції'),
    para('При наївній реалізації: P процесів по черзі надсилають root-у — час O(P). При деревовидній редукції: O(log P). OpenMPI внутрішньо обирає оптимальний алгоритм залежно від розміру повідомлення та топології кластера.'),
  
    h2('Комунікаційні паттерни'),
    bullet('Ring communication: зсув даних по кільцю (Stрічковий алгоритм Фокса).'),
    bullet('Butterfly pattern: паралельна редукція / Allreduce за O(log P) кроків.'),
    bullet('Hypercube: ефективний для алгоритмів з двійковим деревом.'),
  
    h2('Виявлення вузьких місць у MPI-програмах'),
    bullet('mpiP — легковаговий профілювальник MPI: скільки часу витрачено на кожну операцію.'),
    bullet('Intel Trace Analyzer — детальна хронологія MPI-подій, виявлення дисбалансу.'),
    bullet('MPI_Wtime() — точний таймер для вимірювання часових відрізків.'),
    code('double t1 = MPI.Wtime();'),
    code('MPI.COMM_WORLD.Barrier();'),
    code('double elapsed = MPI.Wtime() - t1;'),
  
    h2('Оптимізація доступу до пам\'яті'),
    bullet('Contiguous data: передавати суміжні масиви — MPI може використати zero-copy через RDMA.'),
    bullet('MPI_Type_vector: для рядків двовимірних масивів у мовах з row-major (C) vs column-major (Fortran).'),
    bullet('Numa-aware allocation: виділяти пам\'ять на тому вузлі, де виконується процес.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 44
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 44. Паралельна реалізація алгоритму МГУА'),
  
    h2('Що таке МГУА?'),
    para('Метод Групового Урахування Аргументів (МГУА, GMDH — Group Method of Data Handling) — самоорганізований метод побудови математичних моделей за даними. Розроблений академіком Івахненком (1968). Реалізує нейромережеподібну ієрархічну структуру поліномів.', { justify: true }),
  
    h2('Основна ідея алгоритму МГУА'),
    para('На кожному шарі генерується множина моделей-кандидатів (частіше за все квадратні поліноми від пар аргументів). Відбираються N кращих за критерієм регулярності (похибка на валідаційній вибірці). Вихідні змінні стають вхідними для наступного шару. Процес продовжується, поки похибка зменшується.'),
  
    h2('Шари МГУА та паралелізм'),
    para('Ключова особливість для паралелізації: на кожному шарі всі моделі-кандидати будуються НЕЗАЛЕЖНО одна від одної. Тобто шар є ідеально паралельним — trivially parallelizable.'),
    bullet('Якщо на шарі M моделей-кандидатів, то P потоків обчислюють по M/P моделей.'),
    bullet('Між шарами — синхронізація: відбір N кращих є послідовною операцією.'),
    bullet('Частка паралельного коду: якщо кожен шар обчислює 10^4 поліномів, послідовна фаза відбору — це логарифмічна частина.'),
  
    h2('Реалізація з паралельним шаром на Java'),
    code('ExecutorService exec = Executors.newFixedThreadPool(Runtime.getRuntime().availableProcessors());'),
    code(''),
    code('// Поточний шар — список пар аргументів:'),
    code('List<int[]> pairs = generatePairs(currentInputs.length);'),
    code('List<Future<Model>> futures = new ArrayList<>();'),
    code(''),
    code('for (int[] pair : pairs) {'),
    code('    final int xi = pair[0], xj = pair[1];'),
    code('    futures.add(exec.submit(() -> {'),
    code('        // Побудова квадратного полінома: y = a0 + a1*xi + a2*xj + a3*xi² + a4*xj² + a5*xi*xj'),
    code('        double[] coeffs = leastSquares(xi, xj, trainData);'),
    code('        double error = validationError(coeffs, xi, xj, validData);'),
    code('        return new Model(xi, xj, coeffs, error);'),
    code('    }));'),
    code('}'),
    code(''),
    code('List<Model> models = futures.stream()'),
    code('    .map(f -> { try { return f.get(); } catch (Exception e) { throw new RuntimeException(e); }})'),
    code('    .collect(Collectors.toList());'),
    code(''),
    code('// Відбір N кращих (послідовна фаза):'),
    code('models.sort(Comparator.comparingDouble(m -> m.error));'),
    code('List<Model> selected = models.subList(0, N);'),
  
    h2('Паралелізм з ForkJoin'),
    code('class LayerTask extends RecursiveTask<List<Model>> {'),
    code('    List<int[]> pairs; double[][] data;'),
    code('    protected List<Model> compute() {'),
    code('        if (pairs.size() <= THRESHOLD)'),
    code('            return buildModels(pairs, data);'),
    code('        int mid = pairs.size() / 2;'),
    code('        LayerTask left = new LayerTask(pairs.subList(0, mid), data);'),
    code('        LayerTask right = new LayerTask(pairs.subList(mid, pairs.size()), data);'),
    code('        left.fork();'),
    code('        List<Model> result = right.compute();'),
    code('        result.addAll(left.join());'),
    code('        return result;'),
    code('    }'),
    code('}'),
  
    h2('Оцінка прискорення МГУА'),
    para('Якщо на кожному шарі C(M,2) = M*(M-1)/2 пар, і кожна побудова полінома займає час T_poly, то:'),
    bullet('Послідовний час шару: M*(M-1)/2 * T_poly.'),
    bullet('Паралельний час шару: ceil(M*(M-1)/(2P)) * T_poly.'),
    bullet('Час відбору (послідовний): O(M log M).'),
    bullet('При M=100, P=8: прискорення ~7-8x на фазі побудови.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 46
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 46. Технологія Remote Method Invocation (RMI)'),
  
    h2('Що таке RMI?'),
    para('Java RMI (Remote Method Invocation) — технологія, що дозволяє Java-об\'єктам викликати методи об\'єктів, що знаходяться в іншій JVM (на тому ж або іншому хості). RMI автоматично серіалізує аргументи та повертає результати через мережу.', { justify: true }),
  
    h2('Архітектура RMI'),
    bullet('Remote Interface: інтерфейс, що розширює java.rmi.Remote. Всі методи кидають RemoteException.'),
    bullet('Remote Object (Servant): реалізація Remote Interface. Розширює UnicastRemoteObject.'),
    bullet('RMI Registry: іменований сервіс (на порту 1099 за замовчуванням), де сервер реєструє об\'єкти.'),
    bullet('Stub: проксі на клієнті — виглядає як локальний об\'єкт, але передає виклики по мережі.'),
    bullet('Skeleton (Java 1.x): серверний компонент отримання викликів. У сучасному Java — вбудований у UnicastRemoteObject.'),
  
    h2('Реалізація RMI'),
    para('Крок 1: Remote Interface:'),
    code('import java.rmi.*;'),
    code('public interface MatrixService extends Remote {'),
    code('    double[][] multiply(double[][] A, double[][] B) throws RemoteException;'),
    code('}'),
    para('Крок 2: Сервер:'),
    code('public class MatrixServer extends UnicastRemoteObject implements MatrixService {'),
    code('    public MatrixServer() throws RemoteException {}'),
    code('    public double[][] multiply(double[][] A, double[][] B) throws RemoteException {'),
    code('        return parallelMultiply(A, B);'),
    code('    }'),
    code('    public static void main(String[] args) throws Exception {'),
    code('        Registry reg = LocateRegistry.createRegistry(1099);'),
    code('        reg.bind("MatrixService", new MatrixServer());'),
    code('        System.out.println("Server ready");'),
    code('    }'),
    code('}'),
    para('Крок 3: Клієнт:'),
    code('Registry reg = LocateRegistry.getRegistry("serverHost", 1099);'),
    code('MatrixService service = (MatrixService) reg.lookup("MatrixService");'),
    code('double[][] C = service.multiply(A, B); // Виклик по мережі'),
  
    h2('Серіалізація в RMI'),
    para('Всі аргументи та повернені значення серіалізуються (повинні реалізовувати Serializable). Великі матриці → великий обсяг серіалізації → мережевий bottleneck. Для обчислень з великими даними RMI менш ефективний, ніж MPI або Apache Spark.'),
  
    h2('RMI vs MPI vs REST'),
    bullet('RMI: об\'єктно-орієнтований, тільки Java, синхронний, проксі-модель. Підходить для розподілених Java-систем.'),
    bullet('MPI: мова-незалежний, для HPC, максимальна продуктивність, орієнтований на числові обчислення.'),
    bullet('REST/gRPC: мова-незалежний, HTTP/HTTP2, для мікросервісів та web-API.'),
  
    h2('RMI та Springboot'),
    para('У сучасному Java-розробленні RMI витісняється REST та gRPC. Spring Framework може легко замінити RMI через @RestController або Spring Remoting. SpringBoot дозволяє реалізувати обчислення матриць як REST API — це детальніше у питанні 13 (SpringBoot). RMI залишається актуальним для legacy-систем та Java-специфічних розподілених середовищ.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 48
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 48. Базові складові грід-системи'),
  
    h2('Що таке Grid Computing?'),
    para('Grid (Грід) — це інфраструктура для об\'єднання гетерогенних розподілених обчислювальних ресурсів (процесорів, пам\'яті, дискового простору, наукових інструментів), що знаходяться в різних організаціях та географічних локаціях, у єдиний "суперкомп\'ютер", доступний через мережу.', { justify: true }),
    para('Аналогія: як електрична мережа (grid) надає електроенергію незалежно від її джерела, грід-система надає обчислювальні ресурси незалежно від їх фізичного розташування.'),
  
    h2('Складові грід-системи'),
    para('Ian Foster та Carl Kesselman (The Grid: Blueprint for a New Computing Infrastructure) виділяють п\'ятирівневу архітектуру:'),
  
    bullet('Рівень 1 — Fabrication (Фабрика): фізичні ресурси — обчислювальні вузли, дисковий простір, мережеве обладнання, наукові прилади. Кожен ресурс управляється локально через свою ОС та ПЗ.'),
    bullet('Рівень 2 — Connectivity (З\'єднання): протоколи комунікації та автентифікації. Використовуються стандартні IP-протоколи, SSL/TLS. Єдина аутентифікація: X.509 сертифікати, GSI (Grid Security Infrastructure).'),
    bullet('Рівень 3 — Resource (Ресурс): протоколи та сервіси для доступу до окремих ресурсів. OGSA (Open Grid Services Architecture), WSRF (Web Services Resource Framework), GridFTP для передачі даних.'),
    bullet('Рівень 4 — Collective (Колективний): сервіси що координують множину ресурсів. Реєстр ресурсів (MDS — Monitoring and Discovery System), планувальник задач, брокер ресурсів, сервіс реплікації даних.'),
    bullet('Рівень 5 — Application (Застосунок): кінцеві застосунки та портали користувачів. Grid-застосунки, наукові workflow, портали (myGrid, GENIUS).'),
  
    h2('Ключові компоненти'),
    bullet('Globus Toolkit — де-факто стандарт грід-ПЗ. Надає GSI, GridFTP, GRAM (Grid Resource Allocation Manager), MDS.'),
    bullet('GRAM — менеджер виділення ресурсів та виконання задач. Приймає запити, планує виконання на конкретних вузлах.'),
    bullet('GridFTP — розширення FTP для ефективної передачі великих файлів по WAN. Підтримує паралельні потоки, відновлення після збоїв.'),
    bullet('MDS (Monitoring and Discovery Service) — LDAP-подібний реєстр, де ресурси публікують свої характеристики.'),
    bullet('VOMS (Virtual Organization Membership Service) — управління правами у віртуальних організаціях.'),
  
    h2('Гетерогенність та стандартизація'),
    para('Грід-системи принципово гетерогенні: різні ОС (Linux, Windows, macOS), різні архітектури (x86, ARM, PowerPC), різні локальні планувальники (PBS, SGE, SLURM). Стандартизація через OGSA/Web Services дозволяє абстрагуватись від цих відмінностей.'),
  
    pageBreak()
  );
  
  // ════════════════════════════════════════════════════════════════════════
  // Питання 50
  // ════════════════════════════════════════════════════════════════════════
  sections_content.push(
    h1('Питання 50. Організація і управління розподіленими грід-ресурсами'),
  
    h2('Концепція Virtual Organization (VO)'),
    para('Грід-ресурси належать різним організаціям, але спільно використовуються у рамках Віртуальних Організацій (VO). VO — динамічні групи осіб та інститутів, що поділяють ресурси для досягнення спільних цілей (наприклад, CERN та університети у проекті LHC).', { justify: true }),
    para('Довіра між організаціями забезпечується PKI (Public Key Infrastructure) та X.509 сертифікатами. GSI поширює єдину аутентифікацію без повторного введення паролів (Single Sign-On).'),
  
    h2('Планування задач у грід'),
    para('Грід-планувальник (Grid Scheduler / Broker) вирішує де і коли виконати задачу:'),
    bullet('Збір інформації: запит до MDS про доступні ресурси, їх стан та завантаженість.'),
    bullet('Вибір ресурсу: за критеріями (час виконання, вартість, розташування даних).'),
    bullet('Резервування: запит GRAM на виконання задачі на обраному вузлі.'),
    bullet('Моніторинг: відстеження стану виконання, повторна спроба при збої.'),
    para('Стратегії планування:'),
    bullet('Best-Effort: задача виконується при появі ресурсів.'),
    bullet('Advance Reservation: резервування ресурсів на конкретний час. Важливо для паралельних задач, що потребують одночасного доступу до N вузлів.'),
    bullet('Gang Scheduling: одночасний запуск усіх процесів MPI-задачі.'),
  
    h2('Управління даними в грід'),
    bullet('Реплікація: дані копіюються на близькі до обчислень вузли (Replica Management Service).'),
    bullet('Data Grid: логічна файлова система поверх фізично розподілених сховищ. iRODS, SRM (Storage Resource Manager).'),
    bullet('GridFTP: паралельна передача файлів між вузлами. Підтримує чекпойнти та відновлення.'),
    bullet('Каталоги: метадані про місцезнаходження файлів (LFC — LCG File Catalog).'),
  
    h2('Моніторинг та відмовостійкість'),
    bullet('MDS/BDII (Berkeley Database Information Index) — збір та публікація інформації про стан ресурсів.'),
    bullet('Ganglia, Nagios — моніторинг вузлів у реальному часі.'),
    bullet('Checkpoint/Restart — збереження стану задачі для відновлення при збої вузла.'),
    bullet('Fault Tolerance: автоматичне перезапуск задач на інших вузлах (GRAM підтримує retry-механізм).'),
  
    h2('Приклади грід-систем'),
    bullet('WLCG (Worldwide LHC Computing Grid): обробка даних LHC (CERN). 200 000+ CPU ядер, 200+ PB сховища, 40+ країн.'),
    bullet('EGEE (Enabling Grids for E-sciencE) → EGI: Європейська науково-дослідна грід.'),
    bullet('XSEDE (Extreme Science and Engineering Discovery Environment): американська наукова грід (замінена на ACCESS).'),
    bullet('SETI@home, Folding@home: добровільні грід-системи для наукових обчислень на домашніх ПК.'),
  
    h2('Java та Grid Computing'),
    para('Java відіграє важливу роль у грід-системах:'),
    bullet('Портабельність: "Write Once, Run Anywhere" критично важлива в гетерогенному грід-середовищі.'),
    bullet('JavaGAT (Grid Application Toolkit): Java API для написання переносних грід-застосунків.'),
    bullet('P2P-системи на Java: JXTA, Hazelcast — для розподілених обчислень без централізованого координатора.'),
    bullet('Spring Batch + Cloud: сучасний аналог для batch-обчислень у хмарному середовищі.'),
  
    h2('Від Grid до Cloud'),
    para('Сучасні хмарні системи (AWS, Google Cloud, Azure) реалізують багато концепцій грід: віртуалізація ресурсів, оплата за використання, розподілена робота. Різниця: грід орієнтований на наукові задачі та горизонтальне об\'єднання незалежних організацій; хмара — на комерційний сервіс з централізованим управлінням.'),
  )
  
  // ────────────────────────────────────────────────────────────────────────
  // BUILD DOCUMENT
  // ────────────────────────────────────────────────────────────────────────
  
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 22 }
        }
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 34, bold: true, font: 'Arial', color: '1F3864' },
          paragraph: {
            spacing: { before: 480, after: 240 },
            outlineLevel: 0,
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 6, color: '1F3864', space: 1 }
            }
          }
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '2E5090' },
          paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 1 }
        }
      ]
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: { indent: { left: 720, hanging: 360 } },
                run: { font: 'Arial' }
              }
            }
          ]
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1134, right: 850, bottom: 1134, left: 1418 }
          }
        },
        children: [
          // Title page
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 2000, after: 400 },
            children: [
              new TextRun({
                text: 'Технологія паралельних обчислень',
                bold: true,
                size: 48,
                font: 'Arial',
                color: '1F3864'
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 200 },
            children: [
              new TextRun({
                text: 'Відповіді до екзаменаційних питань',
                size: 32,
                font: 'Arial',
                color: '2E5090'
              })
            ]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 4000 },
            children: [
              new TextRun({
                text: 'Мова реалізації: Java',
                size: 26,
                font: 'Arial',
                italics: true,
                color: '555555'
              })
            ]
          }),
          new Paragraph({ children: [new PageBreak()] }),
          ...sections_content
        ]
      }
    ]
  });
  
  Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('./docs/ev2.docx', buffer);
  }).catch(err => {
    console.error('❌ Помилка:', err);
  });