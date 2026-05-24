// npm install docx
// node generate_exam.js

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, LevelFormat, NumberFormat,
  Header, Footer, TabStopType, TabStopPosition, BorderStyle,
  PageBreak
} = require('docx');
const fs = require('fs');

// ─── helpers ───────────────────────────────────────────────────────────────

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 32, font: 'Arial' })],
    spacing: { before: 400, after: 160 },
    pageBreakBefore: true,
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, font: 'Arial' })],
    spacing: { before: 240, after: 120 },
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, size: 24, font: 'Arial' })],
    spacing: { before: 180, after: 80 },
  });
}

function p(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, font: 'Arial' })],
    spacing: { before: 80, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function bullet(text, ref = 'bullets') {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: [new TextRun({ text, size: 24, font: 'Arial' })],
    spacing: { before: 40, after: 40 },
  });
}

function code(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: 'Courier New', size: 20, color: '2E4057' })],
    spacing: { before: 60, after: 60 },
    indent: { left: 720 },
    border: {
      left: { style: BorderStyle.SINGLE, size: 4, color: '2E75B6', space: 8 },
    },
  });
}

function codeLines(lines) {
  return lines.map(l => code(l));
}

function spacer() {
  return new Paragraph({ children: [new TextRun('')], spacing: { before: 60, after: 60 } });
}

// ─── content ───────────────────────────────────────────────────────────────

const content = [];

// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 1
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 1. Поняття паралельних, псевдопаралельних та розподілених обчислень'));

content.push(h2('1.1. Паралельні обчислення'));
content.push(p('Паралельні обчислення (Parallel Computing) — це модель виконання програми, за якої задача розбивається на підзадачі, що виконуються одночасно на кількох фізичних процесорах або ядрах. Справжній паралелізм можливий лише на багатоядерних або багатопроцесорних системах: кілька потоків (або процесів) дійсно виконуються в один і той самий момент фізичного часу.'));
content.push(p('Мета паралельних обчислень — скорочення часу розв\'язання задачі. Типові приклади: матричне множення, обробка великих масивів даних, машинне навчання, наукові симуляції.'));
content.push(p('Ключові поняття:'));
content.push(bullet('Прискорення (Speedup) S = T1 / Tp, де T1 — час послідовного виконання, Tp — час паралельного на p процесорах.'));
content.push(bullet('Ефективність E = S / p — наскільки добре використовуються ресурси.'));
content.push(bullet('Масштабованість — здатність алгоритму зберігати ефективність при збільшенні кількості процесорів.'));

content.push(h2('1.2. Псевдопаралельні обчислення'));
content.push(p('Псевдопаралельні (конкурентні, concurrent) обчислення — це модель, при якій кілька потоків виконання існують логічно одночасно, проте фізично можуть виконуватись на одному ядрі завдяки швидкому перемиканню контексту (context switching). Операційна система виділяє кожному потоку квант часу (time slice) і по черзі надає йому процесор.'));
content.push(p('В Java псевдопаралелізм є типовою ситуацією при роботі з Thread на однопроцесорній машині. JVM разом з ОС планує потоки. Навіть на багатоядерній машині кількість потоків зазвичай перевищує кількість ядер, тому частина потоків завжди "чекає".'));
content.push(p('Псевдопаралелізм корисний для:'));
content.push(bullet('Підвищення відгуку (responsiveness) застосунку — UI-потік не блокується.'));
content.push(bullet('Перекриття операцій вводу-виводу та обчислень.'));
content.push(bullet('Реалізації серверів з великою кількістю одночасних з\'єднань.'));

content.push(h2('1.3. Розподілені обчислення'));
content.push(p('Розподілені обчислення (Distributed Computing) — це виконання задачі на кількох незалежних комп\'ютерах (вузлах), пов\'язаних мережею. Кожен вузол має свою пам\'ять (модель shared-nothing), а взаємодія відбувається через передачу повідомлень (Message Passing).'));
content.push(p('Стандарти та технології розподілених обчислень:'));
content.push(bullet('MPI (Message Passing Interface) — стандарт для наукових HPC-кластерів (OpenMPI, MPICH).'));
content.push(bullet('Java RMI, gRPC, REST — для розподілених застосунків.'));
content.push(bullet('Apache Spark, Hadoop — для великих даних.'));
content.push(bullet('Spring Boot — мікросервісна архітектура.'));
content.push(p('Основна відмінність від паралельних: вузли фізично розділені, можуть знаходитись в різних датацентрах. Необхідно враховувати затримки мережі, часткові відмови, синхронізацію часу тощо.'));

content.push(h2('1.4. Порівняння моделей'));
content.push(p('Паралельні обчислення: спільна пам\'ять, малі затримки, Thread/Process на одній машині.'));
content.push(p('Псевдопаралельні: одне ядро, перемикання контексту, logically concurrent but physically sequential.'));
content.push(p('Розподілені: окремі машини, передача повідомлень, висока відмовостійкість, горизонтальне масштабування.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 3
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 3. Способи розпаралелювання'));

content.push(h2('3.1. Загальний підхід'));
content.push(p('Розпаралелювання — процес перетворення послідовного алгоритму на паралельний. Головна задача: знайти незалежні частини обчислень, які можна виконувати одночасно. Виділяють кілька підходів залежно від рівня та характеру паралелізму.'));

content.push(h2('3.2. Паралелізм даних (Data Parallelism)'));
content.push(p('Одна операція застосовується до різних підмножин даних. Типовий приклад — стрічковий розподіл рядків матриці між потоками. Кожен потік обробляє свій діапазон рядків незалежно. В Java реалізується через розбиття масиву на частини і запуск потоків:'));
content.push(...codeLines([
  'int rows = n / numThreads;',
  'for (int i = 0; i < numThreads; i++) {',
  '    final int start = i * rows;',
  '    final int end = (i == numThreads-1) ? n : start + rows;',
  '    threads[i] = new Thread(() -> computeRows(start, end));',
  '    threads[i].start();',
  '}',
]));

content.push(h2('3.3. Паралелізм задач (Task Parallelism)'));
content.push(p('Різні задачі (функції) виконуються паралельно. Наприклад, один потік читає дані з мережі, інший — обробляє вже зчитані, третій — записує результат. Це паттерн Producer-Consumer або конвеєр (Pipeline). В Java: ExecutorService, ForkJoinPool, CompletableFuture.'));

content.push(h2('3.4. Розбиття за функціями (Functional Decomposition)'));
content.push(p('Програма розбивається на функціональні блоки, які виконуються паралельно. Підходить для мікросервісів та модульних застосунків. Наприклад, у Spring Boot — окремі сервіси для різних обчислень, що запускаються паралельно через CompletableFuture.'));

content.push(h2('3.5. ForkJoin (рекурсивне розбиття)'));
content.push(p('Задача рекурсивно ділиться на підзадачі до досягнення порогового розміру (threshold). Потім підзадачі вирішуються паралельно, а результати об\'єднуються. Реалізується в Java через ForkJoinPool та RecursiveTask / RecursiveAction:'));
content.push(...codeLines([
  'class SumTask extends RecursiveTask<Long> {',
  '    static final int THRESHOLD = 1000;',
  '    int[] arr; int lo, hi;',
  '    protected Long compute() {',
  '        if (hi - lo <= THRESHOLD) {',
  '            long sum = 0;',
  '            for (int i = lo; i < hi; i++) sum += arr[i];',
  '            return sum;',
  '        }',
  '        int mid = (lo + hi) / 2;',
  '        SumTask left = new SumTask(arr, lo, mid);',
  '        SumTask right = new SumTask(arr, mid, hi);',
  '        left.fork();',
  '        return right.compute() + left.join();',
  '    }',
  '}',
]));

content.push(h2('3.6. Пули потоків (Thread Pool)'));
content.push(p('Замість створення нового потоку для кожної задачі, підтримується фіксований набір (пул) потоків, які отримують задачі з черги. Це усуває overhead на створення/знищення потоків. В Java — Executors.newFixedThreadPool(n), newCachedThreadPool(), newScheduledThreadPool().'));

content.push(h2('3.7. MPI та передача повідомлень'));
content.push(p('Для розподілених систем: кожен процес має свій адресний простір. Взаємодія через MPI_Send/MPI_Recv (точка-точка) або MPI_Bcast, MPI_Scatter, MPI_Gather, MPI_Reduce (колективні). Кожен процес обробляє свій блок даних.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 5
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 5. Способи підвищення продуктивності комп\'ютерів. Суперкомп\'ютери'));

content.push(h2('5.1. Апаратні способи підвищення продуктивності'));
content.push(p('Конвеєризація (Pipelining): розбиття виконання інструкції на стадії (fetch, decode, execute, write-back). Дозволяє починати нову інструкцію, не чекаючи завершення попередньої.'));
content.push(p('Суперскалярні процесори: кілька виконавчих пристроїв, що обробляють декілька інструкцій за один такт. Сучасні CPU Intel/AMD — суперскалярні з Out-of-Order Execution.'));
content.push(p('SIMD (Single Instruction Multiple Data): одна інструкція обробляє вектор даних одночасно (SSE, AVX). Ефективно для матричних операцій та DSP.'));
content.push(p('Кеш-пам\'ять (L1/L2/L3): зменшення часу доступу до даних. Принцип локальності — велика роль для паралельних алгоритмів (cache coherence).'));
content.push(p('Багатоядерні процесори: кілька ядер на одному чіпі зі спільною або роздільною кеш-пам\'яттю.'));
content.push(p('GPGPU: використання тисяч ядер GPU для масивно-паралельних обчислень (CUDA, OpenCL).'));

content.push(h2('5.2. Програмні способи'));
content.push(bullet('Оптимізація алгоритмів — вибір алгоритму з меншою складністю.'));
content.push(bullet('Компіляторна оптимізація — розгортання циклів, векторизація.'));
content.push(bullet('Паралельне програмування — OpenMP, MPI, Java Thread, CUDA.'));
content.push(bullet('Балансування навантаження — рівномірний розподіл задач між ядрами.'));

content.push(h2('5.3. Суперкомп\'ютери'));
content.push(p('Суперкомп\'ютер — обчислювальна система з надзвичайно високою продуктивністю, що вимірюється у флопс (FLOPS — Floating Point Operations Per Second). Сучасні лідери (TOP500) досягають ексафлопс (10^18 FLOPS).'));
content.push(p('Архітектури суперкомп\'ютерів:'));
content.push(bullet('MPP (Massively Parallel Processing) — тисячі вузлів із власною пам\'яттю, з\'єднаних швидкою мережею (InfiniBand). Приклад: Frontier (США, 1.1 ExaFLOPS).'));
content.push(bullet('SMP (Symmetric Multiprocessing) — спільна пам\'ять, кілька процесорів. Менш масштабовані.'));
content.push(bullet('Кластери — стандартні сервери, з\'єднані мережею. Розповсюджений варіант.'));
content.push(bullet('Гібридні — CPU + GPU вузли (Tesla, A100).'));
content.push(p('Програмування суперкомп\'ютерів: MPI для міжвузлової взаємодії, OpenMP для внутрішньовузлового паралелізму, CUDA/ROCm для GPU.'));
content.push(p('Застосування: кліматичне моделювання, ядерні симуляції, геноміка, ШІ, аеродинаміка.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 7
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 7. Поняття потоку обчислень. Створення та запуск потоків в Java'));

content.push(h2('7.1. Поняття потоку'));
content.push(p('Потік (Thread) — найменша одиниця виконання, що планується операційною системою. На відміну від процесу, потоки одного процесу поділяють спільну пам\'ять (heap, статичні поля, відкриті файли), але мають власний стек виклику, лічильник команд та локальні змінні.'));
content.push(p('Переваги потоків над процесами: менший overhead на створення та перемикання, простий обмін даними через спільну пам\'ять. Ризики: race conditions, deadlock, livelock.'));

content.push(h2('7.2. Створення потоків в Java — спосіб 1: успадкування від Thread'));
content.push(...codeLines([
  'class MyThread extends Thread {',
  '    private String name;',
  '    MyThread(String name) { this.name = name; }',
  '',
  '    @Override',
  '    public void run() {',
  '        System.out.println("Потік " + name + " виконується");',
  '    }',
  '}',
  '',
  '// Запуск:',
  'MyThread t = new MyThread("A");',
  't.start(); // НЕ t.run() — start() запускає новий потік!',
]));
content.push(p('Метод start() повідомляє JVM створити новий потік і викликати run() в ньому. Прямий виклик run() виконує код у поточному потоці без паралелізму.'));

content.push(h2('7.3. Створення потоків — спосіб 2: реалізація Runnable'));
content.push(...codeLines([
  'class MyTask implements Runnable {',
  '    @Override',
  '    public void run() {',
  '        System.out.println("Task в потоці: " + Thread.currentThread().getName());',
  '    }',
  '}',
  '',
  '// Запуск:',
  'Thread t = new Thread(new MyTask());',
  't.start();',
  '',
  '// Або через лямбда (Java 8+):',
  'Thread t2 = new Thread(() -> System.out.println("Lambda thread"));',
  't2.start();',
]));
content.push(p('Перевага Runnable: клас може успадковувати інший клас (Java не підтримує множинне успадкування). Runnable — функціональний інтерфейс, тому підтримує лямбди.'));

content.push(h2('7.4. Callable та Future'));
content.push(p('Runnable не повертає результат і не може кидати checked exceptions. Callable<V> вирішує обидві проблеми:'));
content.push(...codeLines([
  'Callable<Integer> task = () -> {',
  '    Thread.sleep(1000);',
  '    return 42;',
  '};',
  'ExecutorService executor = Executors.newSingleThreadExecutor();',
  'Future<Integer> future = executor.submit(task);',
  'Integer result = future.get(); // блокує до завершення',
  'executor.shutdown();',
]));

content.push(h2('7.5. Lifecycle потоку'));
content.push(p('NEW → RUNNABLE → (BLOCKED / WAITING / TIMED_WAITING) → TERMINATED. Стан BLOCKED — потік чекає монітора. WAITING — чекає notify(). TIMED_WAITING — sleep() або wait(timeout).'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 9
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 9. Клас Thread'));

content.push(h2('9.1. Основні методи класу Thread'));
content.push(p('Клас java.lang.Thread є основним засобом керування потоками в Java. Реалізує інтерфейс Runnable.'));
content.push(bullet('start() — запускає новий потік; JVM викликає run() у новому потоці.'));
content.push(bullet('run() — тіло потоку; перевизначається або береться з Runnable.'));
content.push(bullet('sleep(millis) — призупиняє поточний потік на задану кількість мілісекунд (static). Не звільняє монітори.'));
content.push(bullet('join() — поточний потік чекає завершення даного потоку. join(millis) — із таймаутом.'));
content.push(bullet('interrupt() — встановлює прапорець переривання. isInterrupted() — перевірка.'));
content.push(bullet('yield() — підказка планувальнику, що потік готовий поступитись процесором.'));
content.push(bullet('setDaemon(true) — позначає як daemon-потік (JVM завершується, коли залишились тільки daemon-потоки).'));
content.push(bullet('setPriority(int) — пріоритет від MIN_PRIORITY(1) до MAX_PRIORITY(10), за замовчуванням NORM_PRIORITY(5).'));
content.push(bullet('getName() / setName(String) — ім\'я потоку для відлагодження.'));
content.push(bullet('currentThread() — static, повертає посилання на поточний потік.'));
content.push(bullet('getState() — повертає Thread.State (NEW, RUNNABLE, BLOCKED, WAITING, TIMED_WAITING, TERMINATED).'));

content.push(h2('9.2. Пріоритети потоків'));
content.push(p('Пріоритет впливає на частоту надання процесорного часу планувальником ОС, але не є абсолютною гарантією. На різних ОС поведінка може відрізнятись. Не слід покладатись виключно на пріоритети для синхронізації.'));
content.push(...codeLines([
  'Thread high = new Thread(() -> System.out.println("High"));',
  'high.setPriority(Thread.MAX_PRIORITY); // 10',
  'Thread low = new Thread(() -> System.out.println("Low"));',
  'low.setPriority(Thread.MIN_PRIORITY); // 1',
  'high.start();',
  'low.start();',
]));

content.push(h2('9.3. Daemon-потоки'));
content.push(p('Daemon-потоки виконують фонові задачі (GC, таймери). JVM не чекає їх завершення при виході. Важливо: setDaemon() потрібно викликати до start(), інакше IllegalThreadStateException.'));
content.push(...codeLines([
  'Thread daemon = new Thread(() -> {',
  '    while (true) {',
  '        System.out.println("Daemon running...");',
  '        try { Thread.sleep(500); } catch (InterruptedException e) { break; }',
  '    }',
  '});',
  'daemon.setDaemon(true);',
  'daemon.start();',
]));

content.push(h2('9.4. Переривання потоків'));
content.push(p('Правильний спосіб зупинки потоку — переривання (interrupt). Потік перевіряє прапорець і завершує роботу коректно:'));
content.push(...codeLines([
  'Thread worker = new Thread(() -> {',
  '    while (!Thread.currentThread().isInterrupted()) {',
  '        // корисна робота',
  '    }',
  '    System.out.println("Потік завершено коректно");',
  '});',
  'worker.start();',
  'Thread.sleep(1000);',
  'worker.interrupt(); // сигналізуємо про завершення',
]));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 11
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 11. Стрічковий алгоритм паралельного множення матриць'));

content.push(h2('11.1. Постановка задачі'));
content.push(p('Дано матриці A (n×m) та B (m×k). Потрібно обчислити C = A × B, де C[i][j] = Σ A[i][t] × B[t][j] для t від 0 до m-1. Послідовна складність: O(n·m·k). Для квадратних матриць n×n — O(n³).'));

content.push(h2('11.2. Ідея стрічкового алгоритму'));
content.push(p('Матриця A розбивається на горизонтальні "стрічки" (смуги рядків). Кожна стрічка призначається окремому потоку. Матриця B при цьому є спільною і доступна всім потокам (тільки для читання — без race condition). Кожен потік обчислює відповідні рядки результату C.'));
content.push(p('Розбиття: якщо n рядків та p потоків, кожен потік отримує n/p рядків (останній — залишок).'));

content.push(h2('11.3. Реалізація на Java'));
content.push(...codeLines([
  'public class StripedMatrixMultiplication {',
  '    static int[][] A, B, C;',
  '    static int n, m, k;',
  '',
  '    static class WorkerThread extends Thread {',
  '        int startRow, endRow;',
  '        WorkerThread(int s, int e) { startRow=s; endRow=e; }',
  '',
  '        @Override',
  '        public void run() {',
  '            for (int i = startRow; i < endRow; i++) {',
  '                for (int j = 0; j < k; j++) {',
  '                    C[i][j] = 0;',
  '                    for (int t = 0; t < m; t++)',
  '                        C[i][j] += A[i][t] * B[t][j];',
  '                }',
  '            }',
  '        }',
  '    }',
  '',
  '    public static void multiply(int numThreads) throws InterruptedException {',
  '        Thread[] threads = new Thread[numThreads];',
  '        int rowsPerThread = n / numThreads;',
  '        for (int i = 0; i < numThreads; i++) {',
  '            int start = i * rowsPerThread;',
  '            int end = (i == numThreads-1) ? n : start + rowsPerThread;',
  '            threads[i] = new WorkerThread(start, end);',
  '            threads[i].start();',
  '        }',
  '        for (Thread t : threads) t.join();',
  '    }',
  '}',
]));

content.push(h2('11.4. Аналіз'));
content.push(p('Трудомісткість: T(p) ≈ n³ / p (ідеально). Прискорення S = p. Ефективність E = 1 (теоретично). На практиці — втрати на синхронізацію через join(), cache miss при доступі до стовпців B (row-major storage). Для покращення кеш-поведінки транспонують B перед множенням.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 13
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 13. Алгоритм Кеннона паралельного множення матриць'));

content.push(h2('13.1. Призначення алгоритму'));
content.push(p('Алгоритм Кеннона (Cannon\'s Algorithm) — класичний алгоритм паралельного множення квадратних матриць на сітці процесорів p = √P × √P (де P — кількість процесів). Призначений для розподілених систем з передачею повідомлень (MPI). Мінімізує обсяг переданих даних між процесами.'));

content.push(h2('13.2. Ідея алгоритму'));
content.push(p('Матриці A та B розбиваються на блоки, кожен процес (i,j) отримує блок A[i][j] та B[i][j]. Алгоритм виконується за √P кроків:'));
content.push(bullet('Ініціалізація: циклічний зсув рядків A вліво на i позицій, стовпців B вгору на j позицій.'));
content.push(bullet('На кожному кроці: кожен процес множить свої локальні блоки та акумулює в C[i][j].'));
content.push(bullet('Зсув: блоки A зсуваються вліво, блоки B — вгору по тороїдальній сітці.'));
content.push(p('Тороїдальна сітка: процес (i,j) пов\'язаний з сусідами (i, (j+1)%q) та ((i+1)%q, j), де q = √P.'));

content.push(h2('13.3. Переваги'));
content.push(bullet('Рівномірне навантаження — кожен процес виконує однакову кількість роботи.'));
content.push(bullet('Мінімальна комунікація — кожен процес передає лише свій блок сусідам.'));
content.push(bullet('Масштабується на великі матриці при збільшенні кількості процесів.'));

content.push(h2('13.4. Складність'));
content.push(p('Обчислювальна складність: O(n³/P). Комунікаційна: O(n²/√P × √P) = O(n²). Загальне прискорення близьке до лінійного (S ≈ P) при великих матрицях.'));

content.push(h2('13.5. Зв\'язок з алгоритмом Фокса'));
content.push(p('Алгоритм Фокса (Fox\'s Algorithm) подібний, але на кожному кроці процес, що містить діагональний блок A[k][k], розсилає його всім процесам у своєму рядку (broadcast), а не зсуває. Потребує операції колективної комунікації Broadcast. Обидва алгоритми — O(n³/P) за обчисленнями.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 15
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 15. Управління потоками в Java'));

content.push(h2('15.1. Основи управління'));
content.push(p('Управління потоками охоплює: запуск, призупинення, відновлення, пріоритизацію та завершення потоків. Застарілі методи stop(), suspend(), resume() — небезпечні та вилучені. Рекомендований підхід: кооперативне завершення через interrupt() та volatile-прапорці.'));

content.push(h2('15.2. sleep() та yield()'));
content.push(...codeLines([
  '// sleep — призупинення на певний час',
  'try {',
  '    Thread.sleep(2000); // 2 секунди',
  '} catch (InterruptedException e) {',
  '    Thread.currentThread().interrupt(); // відновлюємо прапорець',
  '}',
  '',
  '// yield — підказка планувальнику',
  'Thread.yield(); // Поточний потік готовий поступитись CPU',
]));

content.push(h2('15.3. join()'));
content.push(p('join() — один потік чекає завершення іншого. Критично для збору результатів паралельних обчислень:'));
content.push(...codeLines([
  'Thread[] workers = new Thread[4];',
  'for (int i = 0; i < 4; i++) {',
  '    workers[i] = new Thread(new ComputeTask(i));',
  '    workers[i].start();',
  '}',
  'for (Thread w : workers) {',
  '    try { w.join(); } catch (InterruptedException e) { ... }',
  '}',
  '// Тут всі потоки завершились, результати готові',
]));

content.push(h2('15.4. Volatile та AtomicInteger'));
content.push(p('volatile забезпечує видимість змін між потоками (visibility), але не атомарність. Для атомарних операцій — java.util.concurrent.atomic:'));
content.push(...codeLines([
  'volatile boolean running = true; // visibility гарантована',
  'AtomicInteger counter = new AtomicInteger(0);',
  'counter.incrementAndGet(); // атомарна операція',
]));

content.push(h2('15.5. ThreadGroup'));
content.push(p('ThreadGroup дозволяє групувати потоки та керувати ними разом (наприклад, interrupt() для всієї групи). У сучасному Java рідко використовується — замінений ExecutorService.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 17
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 17. Блокування потоку'));

content.push(h2('17.1. Монітор та synchronized'));
content.push(p('Кожен об\'єкт Java має асоційований монітор (mutex). Блок synchronized захоплює монітор об\'єкта: тільки один потік може виконувати synchronized-код на даному об\'єкті. Інші — блокуються у стані BLOCKED, чекаючи звільнення.'));
content.push(...codeLines([
  'class Counter {',
  '    private int count = 0;',
  '',
  '    // Синхронізований метод — монітор = this',
  '    public synchronized void increment() { count++; }',
  '',
  '    // Синхронізований блок — явний вибір монітора',
  '    public void decrement() {',
  '        synchronized(this) { count--; }',
  '    }',
  '}',
]));

content.push(h2('17.2. Блокування через wait() / notify()'));
content.push(p('Метод wait() змушує потік звільнити монітор та перейти у стан WAITING. notify() / notifyAll() будить потоки, що чекають на даному моніторі. Викликаються тільки у synchronized-контексті:'));
content.push(...codeLines([
  'synchronized(lock) {',
  '    while (!condition) {', // while, не if — захист від spurious wakeup
  '        lock.wait();',
  '    }',
  '    // умова виконана — продовжуємо',
  '}',
  '',
  '// В іншому потоці:',
  'synchronized(lock) {',
  '    condition = true;',
  '    lock.notifyAll();',
  '}',
]));

content.push(h2('17.3. ReentrantLock (java.util.concurrent.locks)'));
content.push(p('Більш гнучка альтернатива synchronized: підтримує tryLock(), lockInterruptibly(), fairness-режим:'));
content.push(...codeLines([
  'ReentrantLock lock = new ReentrantLock(true); // fair lock',
  'lock.lock();',
  'try {',
  '    // критична секція',
  '} finally {',
  '    lock.unlock(); // ЗАВЖДИ у finally!',
  '}',
  '',
  '// Спроба захопити без блокування:',
  'if (lock.tryLock(100, TimeUnit.MILLISECONDS)) {',
  '    try { /* ... */ } finally { lock.unlock(); }',
  '}',
]));

content.push(h2('17.4. ReadWriteLock'));
content.push(p('Якщо читань набагато більше, ніж записів, ReadWriteLock забезпечує паралельне читання та ексклюзивний запис:'));
content.push(...codeLines([
  'ReadWriteLock rwLock = new ReentrantReadWriteLock();',
  '',
  '// Читання — паралельне',
  'rwLock.readLock().lock();',
  'try { return data; } finally { rwLock.readLock().unlock(); }',
  '',
  '// Запис — ексклюзивний',
  'rwLock.writeLock().lock();',
  'try { data = newValue; } finally { rwLock.writeLock().unlock(); }',
]));

content.push(h2('17.5. Deadlock'));
content.push(p('Deadlock — ситуація, коли два потоки чекають монітори один одного. Умови виникнення (Coffman): взаємне виключення, утримання та очікування, відсутність примусового вивільнення, кругове очікування. Запобігання: завжди захоплювати блокування в однаковому порядку, використовувати tryLock з таймаутом.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 19
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 19. Синхронізовані методи'));

content.push(h2('19.1. Ключове слово synchronized'));
content.push(p('Ключове слово synchronized забезпечує взаємне виключення та видимість змін. При синхронізації методу монітором є поточний об\'єкт (this). При синхронізації статичного методу — об\'єкт класу (MyClass.class).'));
content.push(...codeLines([
  'class BankAccount {',
  '    private double balance;',
  '',
  '    public synchronized void deposit(double amount) {',
  '        balance += amount; // атомарно відносно інших synchronized методів',
  '    }',
  '',
  '    public synchronized void withdraw(double amount) {',
  '        if (balance >= amount) balance -= amount;',
  '    }',
  '',
  '    public synchronized double getBalance() { return balance; }',
  '}',
]));

content.push(h2('19.2. Статична синхронізація'));
content.push(...codeLines([
  'class IdGenerator {',
  '    private static int counter = 0;',
  '',
  '    // Монітор = IdGenerator.class',
  '    public static synchronized int nextId() {',
  '        return ++counter;',
  '    }',
  '}',
]));

content.push(h2('19.3. Синхронізований блок'));
content.push(p('Синхронізований блок дозволяє точніше контролювати область блокування (зменшує contention) та вибирати об\'єкт-монітор:'));
content.push(...codeLines([
  'class Repository {',
  '    private final Object readLock = new Object();',
  '    private final Object writeLock = new Object();',
  '    private List<String> data = new ArrayList<>();',
  '',
  '    public void add(String item) {',
  '        synchronized(writeLock) {',
  '            data.add(item); // мінімальна критична секція',
  '        }',
  '    }',
  '',
  '    public int size() {',
  '        synchronized(readLock) {',
  '            return data.size();',
  '        }',
  '    }',
  '}',
]));

content.push(h2('19.4. Reentrancy'));
content.push(p('synchronized в Java — реентрантний: потік може повторно захопити монітор, який вже утримує (наприклад, синхронізований метод викликає інший синхронізований метод того самого об\'єкта). Лічильник входжень відстежується JVM.'));

content.push(h2('19.5. Happened-before'));
content.push(p('Java Memory Model гарантує: всі дії потоку до виходу з synchronized-блоку видимі потоку, що входить у synchronized-блок на тому самому моніторі. Це так звана happens-before гарантія.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 21
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 21. Високорівневі способи управління потоками з використанням java.util.concurrent'));

content.push(h2('21.1. Огляд пакету java.util.concurrent'));
content.push(p('Пакет java.util.concurrent (JUC), введений у Java 5, надає:'));
content.push(bullet('Executor Framework — управління пулами потоків.'));
content.push(bullet('Синхронізатори — CountDownLatch, CyclicBarrier, Semaphore, Phaser.'));
content.push(bullet('Concurrent колекції — ConcurrentHashMap, CopyOnWriteArrayList, LinkedBlockingQueue.'));
content.push(bullet('Атомарні змінні — AtomicInteger, AtomicReference, LongAdder.'));
content.push(bullet('Locks — ReentrantLock, ReadWriteLock, StampedLock.'));
content.push(bullet('CompletableFuture — асинхронне програмування.'));

content.push(h2('21.2. CountDownLatch'));
content.push(p('Дозволяє одному або кільком потокам чекати, поки інші виконають певні операції. Лічильник зменшується через countDown(), потоки чекають через await():'));
content.push(...codeLines([
  'CountDownLatch latch = new CountDownLatch(3);',
  'for (int i = 0; i < 3; i++) {',
  '    new Thread(() -> {',
  '        doWork();',
  '        latch.countDown();',
  '    }).start();',
  '}',
  'latch.await(); // головний потік чекає всіх трьох',
]));

content.push(h2('21.3. CyclicBarrier'));
content.push(p('Бар\'єр, що дозволяє групі потоків чекати одне одного в певній точці. Після досягнення бар\'єру всі потоки продовжують виконання. "Cyclic" — можна перевикористовувати. Використовується в ітеративних паралельних алгоритмах:'));
content.push(...codeLines([
  'CyclicBarrier barrier = new CyclicBarrier(4, () -> System.out.println("Крок завершено"));',
  'for (int i = 0; i < 4; i++) {',
  '    new Thread(() -> {',
  '        for (int step = 0; step < 10; step++) {',
  '            computeStep();',
  '            try { barrier.await(); } catch (Exception e) { ... }',
  '        }',
  '    }).start();',
  '}',
]));

content.push(h2('21.4. Semaphore'));
content.push(p("Контролює доступ до обмеженого ресурсу (наприклад, пул з'єднань до БД). acquire() — зменшує лічильник (блокує якщо 0). release() — збільшує:"));
content.push(...codeLines([
  'Semaphore semaphore = new Semaphore(5); // max 5 одночасних доступів',
  'semaphore.acquire();',
  'try { useResource(); }',
  'finally { semaphore.release(); }',
]));

content.push(h2('21.5. CompletableFuture'));
content.push(p('Потужний API для асинхронного програмування та композиції асинхронних задач:'));
content.push(...codeLines([
  'CompletableFuture<Integer> future = CompletableFuture',
  '    .supplyAsync(() -> fetchData())         // async у ForkJoinPool',
  '    .thenApply(data -> process(data))       // трансформація',
  '    .thenCombine(otherFuture, (a,b) -> a+b) // комбінування',
  '    .exceptionally(ex -> defaultValue);     // обробка помилок',
  'future.thenAccept(System.out::println);',
]));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 23
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 23. Пули потоків'));

content.push(h2('23.1. Навіщо потрібні пули'));
content.push(p('Створення та знищення потоків — дорогі операції (виділення стеку, реєстрація в ОС). При великій кількості короткочасних задач (наприклад, HTTP-запити в сервері) overhead суттєвий. Пул потоків: заздалегідь створені потоки, що очікують задач у черзі. Задача виконується існуючим потоком, після чого потік повертається в пул.'));

content.push(h2('23.2. Executor Framework'));
content.push(p('Інтерфейс Executor: execute(Runnable). ExecutorService: розширює, додає submit(), shutdown(), invokeAll(). ThreadPoolExecutor — базова реалізація.'));

content.push(h2('23.3. Типи пулів через Executors'));
content.push(bullet('newFixedThreadPool(n) — фіксована кількість потоків, задачі в unbounded черзі.'));
content.push(bullet('newCachedThreadPool() — потоки створюються за потребою, не використовувані видаляються через 60 с. Для коротких задач.'));
content.push(bullet('newSingleThreadExecutor() — один потік, гарантує послідовне виконання.'));
content.push(bullet('newScheduledThreadPool(n) — для задач з затримкою або виконання за розкладом.'));
content.push(bullet('newWorkStealingPool() — ForkJoinPool з parallelism = кількість ядер. Work-stealing.'));

content.push(h2('23.4. Приклад — FixedThreadPool для системи масового обслуговування'));
content.push(...codeLines([
  'ExecutorService pool = Executors.newFixedThreadPool(4); // 4 "сервери"',
  'BlockingQueue<Runnable> taskQueue = new LinkedBlockingQueue<>();',
  '',
  '// Генератор заявок',
  'ScheduledExecutorService generator = Executors.newScheduledThreadPool(1);',
  'generator.scheduleAtFixedRate(() -> {',
  '    pool.submit(() -> {',
  '        System.out.println("Обробка заявки потоком: " + Thread.currentThread().getName());',
  '        try { Thread.sleep(ThreadLocalRandom.current().nextInt(100, 500)); }',
  '        catch (InterruptedException e) { Thread.currentThread().interrupt(); }',
  '    });',
  '}, 0, 200, TimeUnit.MILLISECONDS); // нова заявка кожні 200 мс',
  '',
  'Thread.sleep(5000);',
  'generator.shutdown();',
  'pool.shutdown();',
  'pool.awaitTermination(10, TimeUnit.SECONDS);',
]));

content.push(h2('23.5. Налаштування ThreadPoolExecutor'));
content.push(p('Для точного контролю використовують конструктор ThreadPoolExecutor(corePoolSize, maximumPoolSize, keepAliveTime, unit, workQueue, threadFactory, rejectedExecutionHandler). RejectedExecutionHandler — політика при переповненні черги: AbortPolicy (виняток), CallerRunsPolicy (виконує caller), DiscardPolicy, DiscardOldestPolicy.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 25
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 25. Розробка паралельних програм з використанням ForkJoinFramework'));

content.push(h2('25.1. Ідея ForkJoin'));
content.push(p('ForkJoin Framework (Java 7+) реалізує парадигму "divide and conquer". Задача рекурсивно розбивається (fork) на підзадачі до порогового розміру, вирішується послідовно, результати об\'єднуються (join). Базується на ForkJoinPool з work-stealing алгоритмом: незайнятий потік "краде" задачі з черги зайнятого.'));

content.push(h2('25.2. RecursiveTask та RecursiveAction'));
content.push(p('RecursiveTask<V> — повертає результат. RecursiveAction — не повертає (side effects). Обидва є підкласами ForkJoinTask.'));
content.push(...codeLines([
  '// Пошук у файловій системі — RecursiveTask',
  'class FileSearchTask extends RecursiveTask<List<Path>> {',
  '    private final Path dir;',
  '    private final String pattern;',
  '    static final int THRESHOLD = 10; // <= 10 файлів — послідовно',
  '',
  '    FileSearchTask(Path dir, String pattern) {',
  '        this.dir = dir; this.pattern = pattern;',
  '    }',
  '',
  '    @Override',
  '    protected List<Path> compute() {',
  '        List<Path> result = new ArrayList<>();',
  '        List<FileSearchTask> subTasks = new ArrayList<>();',
  '',
  '        try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir)) {',
  '            List<Path> entries = new ArrayList<>();',
  '            for (Path entry : stream) entries.add(entry);',
  '',
  '            if (entries.size() <= THRESHOLD) {',
  '                // Послідовний пошук',
  '                for (Path entry : entries) {',
  '                    if (Files.isDirectory(entry)) {',
  '                        result.addAll(new FileSearchTask(entry, pattern).compute());',
  '                    } else if (entry.getFileName().toString().contains(pattern)) {',
  '                        result.add(entry);',
  '                    }',
  '                }',
  '            } else {',
  '                // Паралельний поділ',
  '                for (Path entry : entries) {',
  '                    if (Files.isDirectory(entry)) {',
  '                        FileSearchTask task = new FileSearchTask(entry, pattern);',
  '                        task.fork(); // запуск паралельно',
  '                        subTasks.add(task);',
  '                    } else if (entry.getFileName().toString().contains(pattern)) {',
  '                        result.add(entry);',
  '                    }',
  '                }',
  '                for (FileSearchTask t : subTasks)',
  '                    result.addAll(t.join()); // збір результатів',
  '            }',
  '        } catch (IOException e) { e.printStackTrace(); }',
  '        return result;',
  '    }',
  '}',
  '',
  '// Використання:',
  'ForkJoinPool pool = new ForkJoinPool();',
  'List<Path> found = pool.invoke(new FileSearchTask(Paths.get("/data"), ".log"));',
]));

content.push(h2('25.3. Work-Stealing'));
content.push(p('Кожен потік ForkJoinPool має власну deque (двосторонню чергу). Новий fork() додає задачу до хвосту власної черги. Незайнятий потік бере задачі з голови чужої черги (stealing). Це мінімізує блокування та балансує навантаження без централізованої черги.'));

content.push(h2('25.4. Коли використовувати'));
content.push(bullet('Задачі, що легко рекурсивно діляться (сортування злиттям, швидке сортування).'));
content.push(bullet('Пошук у деревах та графах.'));
content.push(bullet('Паралельна обробка колекцій (parallelStream() всередині використовує ForkJoinPool.commonPool()).'));
content.push(bullet('Матричне множення блочним методом.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 27
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 27. Модель обчислень у вигляді графа «операції—операнди»'));

content.push(h2('27.1. Визначення'));
content.push(p('Граф "операції—операнди" (Operation-Operand Graph, також DAG — Directed Acyclic Graph) — це математична модель програми, що відображає залежності між обчисленнями. Вузли графа — операції або операнди (дані). Ребра — потоки даних: ребро від A до B означає, що B використовує результат A.'));

content.push(h2('27.2. Призначення'));
content.push(p('Граф використовується для:'));
content.push(bullet('Виявлення паралелізму: вузли без залежностей можна виконувати паралельно.'));
content.push(bullet('Планування (scheduling): визначення оптимального порядку виконання.'));
content.push(bullet('Оцінки критичного шляху: найдовший шлях у DAG — нижня межа часу паралельного виконання.'));
content.push(bullet('Компіляторної оптимізації: переупорядкування та векторизація.'));

content.push(h2('27.3. Критичний шлях'));
content.push(p('Критичний шлях — найдовша (за часом) послідовність залежних операцій від входу до виходу. Час паралельного виконання T∞ ≥ довжина критичного шляху. Ступінь паралелізму = T1 / T∞, де T1 — час послідовного виконання.'));

content.push(h2('27.4. Приклад'));
content.push(p('Обчислення (a + b) × (c + d): операції s1 = a+b та s2 = c+d незалежні (можуть виконуватись паралельно), потім result = s1 × s2. Критичний шлях: 2 операції замість 3 послідовних.'));

content.push(h2('27.5. Зв\'язок з Java'));
content.push(p('CompletableFuture явно відображає граф залежностей в коді: thenApply, thenCombine, thenCompose утворюють DAG асинхронних обчислень. ForkJoinTask також утворює дерево (граф) підзадач.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 29
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 29. Показники ефективності паралельного алгоритму'));

content.push(h2('29.1. Прискорення (Speedup)'));
content.push(p('S(p) = T(1) / T(p), де T(1) — час найкращого послідовного алгоритму, T(p) — час паралельного на p процесорах. Лінійне прискорення S(p) = p — ідеал. Суперлінійне S(p) > p — можливе через кеш-ефекти. S(p) < p — типова ситуація через комунікацію та синхронізацію.'));

content.push(h2('29.2. Ефективність (Efficiency)'));
content.push(p('E(p) = S(p) / p = T(1) / (p × T(p)). Показує, наскільки добре використовується кожен процесор. E = 1 — ідеальне використання. E < 1 — простої через синхронізацію, комунікацію, незбалансованість. Мета: E > 0.7–0.8 на практиці.'));

content.push(h2('29.3. Накладні витрати'));
content.push(p('Overhead паралельного алгоритму: T_overhead = p × T(p) — T(1). Включає: час синхронізації, передачу повідомлень, незбалансованість, зайву роботу.'));

content.push(h2('29.4. Масштабованість'));
content.push(p('Сильна масштабованість (Strong scaling): S(p) зростає при фіксованому розмірі задачі. Слабка масштабованість (Weak scaling): E залишається постійним при збільшенні p та n одночасно (задача на процесор стала). Weak scaling важлива для великих HPC задач.'));

content.push(h2('29.5. Ізоефективність'));
content.push(p('Функція ізоефективності W(p) — мінімальний розмір задачі, при якому ефективність залишається постійною при збільшенні p. Чим повільніше зростає W(p), тим краще масштабується алгоритм.'));

content.push(h2('29.6. Granularity (зернистість)'));
content.push(p('Fine-grained parallelism — багато малих задач, великий overhead комунікації. Coarse-grained — мало великих задач, менший overhead але гірший баланс. Оптимальна зернистість залежить від задачі та архітектури.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 31
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 31. Закон Амдала. Ефект Амдала'));

content.push(h2('31.1. Формулювання закону'));
content.push(p('Закон Амдала (Gene Amdahl, 1967) — фундаментальне обмеження прискорення паралельної програми. Нехай f — частка програми, що не може бути розпаралелена (послідовна частина). Тоді максимальне теоретичне прискорення:'));
content.push(p('S(p) = 1 / (f + (1-f)/p)'));
content.push(p('При p → ∞:  S_max = 1/f'));
content.push(p('Тобто якщо 10% програми послідовні, максимальне прискорення — 10× (незалежно від кількості процесорів).'));

content.push(h2('31.2. Наслідки'));
content.push(bullet('f = 0.5 (50% послідовний): S_max = 2× навіть з нескінченною кількістю процесорів.'));
content.push(bullet('f = 0.1 (10% послідовний): S_max = 10×.'));
content.push(bullet('f = 0.01 (1% послідовний): S_max = 100×.'));
content.push(p('Закон пояснює, чому нескінченне збільшення кількості процесорів не дає нескінченного прискорення. Найважливіший висновок: мінімізація послідовної частини критично важлива.'));

content.push(h2('31.3. Ефект Амдала'));
content.push(p('Практичний висновок: при фіксованому розмірі задачі прискорення швидко насичується. Після певного числа процесорів додавання нових не дає відчутного виграшу, але збільшує витрати (вартість обладнання, енергія). Наприклад, при f = 0.1: 8 проц. → S ≈ 4.7, 64 проц. → S ≈ 8.7, 512 проц. → S ≈ 9.8 (насичення поблизу 10).'));

content.push(h2('31.4. Закон Густафсона'));
content.push(p('Альтернатива Амдалу: при weak scaling (розмір задачі зростає з p) прискорення S(p) = p - f×(p-1) — майже лінійне. Практично важливіше для великих HPC задач: збільшуємо і задачу, і кількість процесорів.'));

content.push(h2('31.5. Приклад розрахунку'));
content.push(p('Задача: T1 = 100 с, f = 0.2 (20% послідовно), p = 4 процесори.'));
content.push(p('T(p) = f × T1 + (1-f) × T1 / p = 0.2×100 + 0.8×100/4 = 20 + 20 = 40 с.'));
content.push(p('S(4) = 100/40 = 2.5. E = 2.5/4 = 62.5%.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 33
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 33. Оцінка ефективності паралельних обчислень'));

content.push(h2('33.1. Методологія оцінки'));
content.push(p('Оцінка ефективності включає теоретичний аналіз та практичне вимірювання (benchmarking). Теоретичний аналіз базується на моделях (Amdahl, PRAM, BSP). Практичне вимірювання — реальні часи виконання на конкретному обладнанні.'));

content.push(h2('33.2. Вимірювання часу в Java'));
content.push(...codeLines([
  'long startTime = System.nanoTime();',
  '// паралельні обчислення',
  'long endTime = System.nanoTime();',
  'double elapsedMs = (endTime - startTime) / 1_000_000.0;',
  'System.out.printf("Час виконання: %.2f мс%n", elapsedMs);',
]));

content.push(h2('33.3. Джерела неефективності'));
content.push(bullet('Комунікаційні затримки (latency, bandwidth) при передачі даних між потоками/процесами.'));
content.push(bullet('Синхронізаційні бар\'єри — потоки чекають найповільнішого.'));
content.push(bullet('Незбалансованість навантаження (load imbalance) — частина процесорів простоює.'));
content.push(bullet('Cache coherence overhead — протоколи MESI/MOESI для підтримки узгодженості кешу.'));
content.push(bullet('False sharing — два потоки змінюють різні змінні в одній кеш-лінії (64 байти).'));

content.push(h2('33.4. False Sharing'));
content.push(p('Критична проблема: різні потоки модифікують сусідні елементи масиву, що потрапляють в одну кеш-лінію. Це призводить до постійної інвалідації кешу між ядрами. Рішення: @Contended (JVM), паддінг масивів:'));
content.push(...codeLines([
  '// ПОГАНО — false sharing',
  'int[] results = new int[4]; // всі 4 int в одній кеш-лінії',
  '',
  '// ДОБРЕ — розміщення з відступом',
  'long[][] padded = new long[4][8]; // кожен елемент [i][0] в окремій кеш-лінії',
]));

content.push(h2('33.5. Профілювання'));
content.push(p('Інструменти: Java Flight Recorder (JFR), JProfiler, VisualVM, async-profiler. Дозволяють знайти вузькі місця: contended locks, GC pauses, thread wait times.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 35
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 35. Архітектура розподілених систем. Таксономія Фліна'));

content.push(h2('35.1. Таксономія Фліна (Flynn\'s Taxonomy)'));
content.push(p('Запропонована Michael Flynn у 1966 р. Класифікує комп\'ютерні архітектури за потоками інструкцій та даних:'));
content.push(bullet('SISD (Single Instruction, Single Data) — класичний фон Нейман, послідовний процесор.'));
content.push(bullet('SIMD (Single Instruction, Multiple Data) — векторні процесори, GPU, SSE/AVX. Одна інструкція обробляє кілька елементів даних.'));
content.push(bullet('MISD (Multiple Instruction, Single Data) — рідкісна, застосовується у відмовостійких системах (різні алгоритми на тих самих даних для перевірки).'));
content.push(bullet('MIMD (Multiple Instruction, Multiple Data) — сучасні багатопроцесорні системи, кластери, суперкомп\'ютери. Кожен процесор виконує свою програму на своїх даних.'));

content.push(h2('35.2. Підкатегорії MIMD'));
content.push(p('Shared Memory MIMD:'));
content.push(bullet('UMA (Uniform Memory Access) — однорідний доступ до пам\'яті. SMP-системи.'));
content.push(bullet('NUMA (Non-Uniform Memory Access) — доступ до локальної пам\'яті швидший. Сучасні багатосокетні сервери AMD EPYC.'));
content.push(p('Distributed Memory MIMD:'));
content.push(bullet('MPP (Massively Parallel Processors) — тисячі вузлів, кожен зі своєю пам\'яттю та процесором.'));
content.push(bullet('Кластери — commodity hardware, з\'єднані мережею (InfiniBand, Ethernet).'));

content.push(h2('35.3. Архітектура розподілених систем'));
content.push(p('Розподілена система: набір незалежних комп\'ютерів, що представляється користувачам єдиним. Ключові характеристики: прозорість (transparency), відмовостійкість (fault tolerance), масштабованість.'));
content.push(p('Архітектурні стилі:'));
content.push(bullet('Клієнт-сервер: centralized server, multiple clients.'));
content.push(bullet('P2P: рівноправні вузли (BitTorrent, blockchain).'));
content.push(bullet('Мікросервіси: незалежні сервіси зі своїми БД, взаємодія через API.'));
content.push(bullet('Middleware: CORBA, Java EE, Spring Boot для абстракції комунікації.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 37
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 37. Стандарти та технології обміну повідомленнями в розподілених системах'));

content.push(h2('37.1. MPI (Message Passing Interface)'));
content.push(p('MPI — стандарт передачі повідомлень для HPC (1994, MPI-1; 2012, MPI-3). Реалізації: OpenMPI, MPICH, Intel MPI. Забезпечує: портабельність, ефективність, повний набір комунікаційних примітивів. Основа для наукових HPC-програм.'));

content.push(h2('37.2. Java RMI (Remote Method Invocation)'));
content.push(p('RMI дозволяє викликати методи об\'єктів на інших JVM прозоро. Stub на клієнті серіалізує аргументи, передає мережею, skeleton на сервері десеріалізує і викликає метод. Обмеження: тільки Java, відносно повільно.'));

content.push(h2('37.3. REST / HTTP'));
content.push(p('Найпоширеніший підхід для мікросервісів. HTTP-методи GET/POST/PUT/DELETE. JSON/XML серіалізація. Stateless. Spring Boot автоматизує розробку REST API. Низька ефективність для HPC через overhead HTTP.'));

content.push(h2('37.4. gRPC'));
content.push(p('Google Remote Procedure Call. Protocol Buffers (protobuf) для серіалізації — компактніший та швидший за JSON. HTTP/2 транспорт. Streaming (server-side, client-side, bidirectional). Підходить для мікросервісів з вимогами до продуктивності.'));

content.push(h2('37.5. Message Brokers'));
content.push(p('Apache Kafka, RabbitMQ — асинхронний обмін повідомленнями. Розв\'язання виробника та споживача. Висока відмовостійкість та масштабованість. Паттерн publish-subscribe.'));

content.push(h2('37.6. Порівняльна таблиця'));
content.push(p('MPI: максимальна продуктивність, тільки для HPC/кластерів. REST: простота, universal, для web. gRPC: ефективний для мікросервісів. RMI: тільки Java. Kafka: великий потік подій, async.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 39
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 39. Методи обміну повідомленнями стандарту MPI один до одного'));

content.push(h2('39.1. Точкова комунікація в MPI'));
content.push(p('MPI Point-to-Point комунікація: один процес відправляє (sender), інший отримує (receiver). Ідентифікація: rank (номер процесу в комунікаторі) та tag (мітка повідомлення).'));

content.push(h2('39.2. Блокуючі операції'));
content.push(p('MPI_Send(buf, count, datatype, dest, tag, comm) — блокує відправника до отримання даних у буфер МПА (або до завершення передачі залежно від реалізації).'));
content.push(p('MPI_Recv(buf, count, datatype, source, tag, comm, status) — блокує отримувача до повного отримання повідомлення.'));
content.push(p('Небезпека: deadlock якщо обидва процеси чекають один одного. Розв\'язання: чергувати Send і Recv, або використовувати MPI_Sendrecv.'));

content.push(h2('39.3. Неблокуючі операції'));
content.push(p('MPI_Isend та MPI_Irecv повертають одразу, надаючи об\'єкт MPI_Request. Перекриває комунікацію та обчислення (communication hiding). MPI_Wait або MPI_Test перевіряє завершення:'));
content.push(...codeLines([
  'MPI_Request req;',
  'MPI_Isend(buffer, count, MPI_INT, dest, tag, MPI_COMM_WORLD, &req);',
  '// Обчислення, поки дані передаються...',
  'do_computation();',
  'MPI_Wait(&req, MPI_STATUS_IGNORE); // чекаємо завершення передачі',
]));

content.push(h2('39.4. Варіанти Send'));
content.push(bullet('MPI_Ssend (Synchronous) — блокує до підтвердження отримання від receiver.'));
content.push(bullet('MPI_Bsend (Buffered) — копіює у локальний буфер, не чекає receiver.'));
content.push(bullet('MPI_Rsend (Ready) — відправник знає, що receiver вже викликав MPI_Recv.'));

content.push(h2('39.5. Застосування для матричного множення'));
content.push(p('Стрічковий розподіл з MPI: процес 0 (master) розсилає стрічки матриці A іншим процесам (MPI_Send у циклі), всі процеси мають повну матрицю B (MPI_Bcast або MPI_Send кожному). Кожен обчислює свої рядки C. Результат збирається через MPI_Recv на процесі 0 або MPI_Gather.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 41
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 41. OpenMPI: особливості побудови програми та основні методи'));

content.push(h2('41.1. Структура MPI-програми'));
content.push(p('Кожна MPI-програма запускається як p незалежних процесів. Всі виконують один і той самий код, але гілкуються за rank. Основна структура:'));
content.push(...codeLines([
  '#include <mpi.h>',
  '#include <stdio.h>',
  '',
  'int main(int argc, char* argv[]) {',
  '    MPI_Init(&argc, &argv);         // ініціалізація MPI',
  '',
  '    int rank, size;',
  '    MPI_Comm_rank(MPI_COMM_WORLD, &rank); // номер поточного процесу',
  '    MPI_Comm_size(MPI_COMM_WORLD, &size); // загальна кількість процесів',
  '',
  '    if (rank == 0) {',
  '        printf("Master process, total: %d\\n", size);',
  '    } else {',
  '        printf("Worker %d\\n", rank);',
  '    }',
  '',
  '    MPI_Finalize();                 // завершення MPI',
  '    return 0;',
  '}',
]));
content.push(p('Компіляція: mpicc -o program program.c. Запуск: mpirun -np 4 ./program'));

content.push(h2('41.2. Колективні операції'));
content.push(bullet('MPI_Bcast(buf, count, type, root, comm) — розсилка від root всім процесам.'));
content.push(bullet('MPI_Scatter(sendbuf, sendcount, sendtype, recvbuf, recvcount, recvtype, root, comm) — розподіл різних частин масиву між процесами.'));
content.push(bullet('MPI_Gather — зворотнє до Scatter: збирає частини від всіх на root.'));
content.push(bullet('MPI_Allgather — кожен процес отримує дані від усіх.'));
content.push(bullet('MPI_Reduce(sendbuf, recvbuf, count, type, op, root, comm) — редукція (MPI_SUM, MPI_MAX, MPI_MIN, MPI_PROD) на root.'));
content.push(bullet('MPI_Allreduce — результат редукції отримують всі.'));
content.push(bullet('MPI_Barrier(comm) — синхронізаційний бар\'єр: чекає всіх.'));

content.push(h2('41.3. Типи даних MPI'));
content.push(p('MPI_INT, MPI_DOUBLE, MPI_CHAR, MPI_FLOAT — відповідники C-типів. MPI_BYTE — сирі байти. Можна визначати похідні типи (MPI_Type_contiguous, MPI_Type_vector) для нерегулярних структур даних (наприклад, стовпців матриці).'));

content.push(h2('41.4. Комунікатори'));
content.push(p('MPI_COMM_WORLD — всі процеси. MPI_Comm_split та MPI_Comm_create дозволяють створювати підгрупи процесів. Декартові топології (MPI_Cart_create) для організації процесів у 2D/3D сітки (для алгоритму Кеннона).'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 43
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 43. Алгоритми паралельного сортування'));

content.push(h2('43.1. Паралельне сортування злиттям'));
content.push(p('Масив ділиться на рівні частини між процесами/потоками. Кожен сортує свою частину локально. Потім попарне злиття відсортованих частин відбувається паралельно (butterfly pattern). Складність: O(n/p × log(n/p)) + O(log p) рівнів злиття. У Java з ForkJoinPool:'));
content.push(...codeLines([
  'class ParallelMergeSort extends RecursiveAction {',
  '    int[] arr; int lo, hi;',
  '    static final int THRESHOLD = 1000;',
  '',
  '    protected void compute() {',
  '        if (hi - lo <= THRESHOLD) {',
  '            Arrays.sort(arr, lo, hi);',
  '            return;',
  '        }',
  '        int mid = (lo + hi) / 2;',
  '        ParallelMergeSort left = new ParallelMergeSort(arr, lo, mid);',
  '        ParallelMergeSort right = new ParallelMergeSort(arr, mid, hi);',
  '        invokeAll(left, right);',
  '        merge(arr, lo, mid, hi);',
  '    }',
  '}',
  'Arrays.parallelSort(arr); // вбудоване паралельне сортування в Java 8+',
]));

content.push(h2('43.2. Quicksort з паралелізмом'));
content.push(p('Після вибору pivot та поділу масиву, ліва та права частини сортуються паралельно. ForkJoinPool ефективно підтримує цей рекурсивний підхід.'));

content.push(h2('43.3. Bitonic Sort'));
content.push(p('Мережа сортування (sorting network) — алгоритм компаратор-базованого сортування, де послідовність порівнянь фіксована незалежно від даних. Ефективний для паралельного hardware. Складність: O(log²n) паралельних кроків.'));

content.push(h2('43.4. Odd-Even Transposition Sort'));
content.push(p('Для MPI: процеси обмінюються між сусідами на парних/непарних ітераціях. Простий але не найефективніший: O(p) ітерацій. Підходить для вже майже відсортованих даних.'));

content.push(h2('43.5. Sample Sort'));
content.push(p('Масштабований алгоритм для великих розподілених систем: sample вибірка → визначення splitters (розділювачів) → перерозподіл за splitters → локальне сортування. Використовується в Hadoop/Spark.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 45
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 45. Оцінювання продуктивності паралельних алгоритмів в розподілених системах'));

content.push(h2('45.1. Специфіка розподілених систем'));
content.push(p('На відміну від систем зі спільною пам\'яттю, розподілені системи мають явні комунікаційні витрати. Модель продуктивності включає: T_comp — час обчислень, T_comm — час комунікацій. T_total = T_comp + T_comm.'));

content.push(h2('45.2. Модель α-β'));
content.push(p('Час передачі повідомлення розміром m байт: T = α + β×m. α (latency) — затримка на встановлення з\'єднання (незалежно від розміру). β (bandwidth) — час передачі одного байта = 1/bandwidth. Для малих повідомлень домінує α, для великих — β×m.'));

content.push(h2('45.3. BSP-модель'));
content.push(p('Bulk Synchronous Parallel: обчислення відбуваються супersteps. Кожен superstep: локальні обчислення → комунікація → глобальний бар\'єр. Продуктивність: T = число_supersteps × (max_обчислень + g×max_повідомлень + l), де g — bandwidth gap, l — latency.'));

content.push(h2('45.4. Bandwidth та Latency вимірювання'));
content.push(p('Практичне вимірювання: ping-pong тест між двома процесами. MPI_Wtime() — точний таймер. Latency: час roundtrip малого повідомлення / 2. Bandwidth: розмір / час для великого повідомлення.'));

content.push(h2('45.5. Комунікаційна ефективність'));
content.push(p('Відношення часу обчислень до часу комунікацій визначає ефективність. Правило: T_comp >> T_comm для гарної масштабованості. Для матричного множення n×n на p процесах: T_comp = O(n³/p), T_comm = O(n²/√p) — при великих n ефективно.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 47
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 47. Грід-технології. Поняття спільного віртуального ресурсу'));

content.push(h2('47.1. Визначення Grid'));
content.push(p('Грід (Grid Computing) — інфраструктура для координованого використання гетерогенних розподілених ресурсів (обчислювальна потужність, дискове сховище, спеціалізоване обладнання), що належать різним організаціям. Термін запропонований Ян Фостером (Ian Foster) за аналогією з електромережею (electrical grid).'));

content.push(h2('47.2. Спільний віртуальний ресурс (Virtual Shared Resource)'));
content.push(p('Ключова концепція Grid: ресурси (CPU часи, диски, прилади) різних організацій об\'єднуються у віртуальний спільний ресурс. Користувач бачить єдине обчислювальне середовище незалежно від фізичного розміщення ресурсів. Організації, що надають ресурси — Virtual Organization (VO).'));

content.push(h2('47.3. Характеристики Grid'));
content.push(bullet('Гетерогенність: різне апаратне та програмне забезпечення.'));
content.push(bullet('Географічна розподіленість: вузли по всьому світу.'));
content.push(bullet('Динамічність: вузли приєднуються та відключаються.'));
content.push(bullet('Автономність: кожна організація контролює свої ресурси.'));
content.push(bullet('Захищеність: аутентифікація (PKI, X.509 сертифікати).'));

content.push(h2('47.4. Типи Grid'));
content.push(bullet('Обчислювальний Grid (Computational Grid): надання CPU. WLCG (Worldwide LHC Computing Grid) для обробки даних CERN.'));
content.push(bullet('Сховищний Grid (Data Grid): розподілене зберігання та доступ до даних.'));
content.push(bullet('Сервісний Grid (Service Grid): надання програмних сервісів.'));
content.push(bullet('Сенсорний Grid: об\'єднання вимірювальних приладів.'));

content.push(h2('47.5. Grid vs Cloud'));
content.push(p('Cloud Computing — комерційна модель надання ресурсів за запитом (AWS, GCP, Azure). Grid — федеративна некомерційна модель для наукових установ. Сучасні наукові Grid часто використовують хмарну інфраструктуру.'));


// ═══════════════════════════════════════════════════════════════════
// ПИТАННЯ 49
// ═══════════════════════════════════════════════════════════════════
content.push(h1('Питання 49. Програмне забезпечення грід-технологій'));

content.push(h2('49.1. Globus Toolkit'));
content.push(p('Globus Toolkit (GT) — де-факто стандарт ПЗ для Grid (Аргоннська національна лабораторія + Чиказький університет). Компоненти:'));
content.push(bullet('GRAM (Grid Resource Allocation and Management) — запуск задач на вузлах Grid.'));
content.push(bullet('GridFTP — розширений FTP для високошвидкісної передачі великих файлів (паралельні потоки).'));
content.push(bullet('GSI (Grid Security Infrastructure) — PKI аутентифікація та авторизація.'));
content.push(bullet('MDS (Monitoring and Discovery Service) — реєстрація та пошук ресурсів.'));

content.push(h2('49.2. OGSA та OGSI'));
content.push(p('OGSA (Open Grid Services Architecture) — архітектурна специфікація на основі Web Services. Grid-ресурс представляється як Grid Service зі стандартним інтерфейсом. WSRF (Web Services Resource Framework) — конкретизація для stateful web services.'));

content.push(h2('49.3. BOINC'));
content.push(p('Berkeley Open Infrastructure for Network Computing — ПЗ для добровільних Grid (volunteer computing). SETI@home, Folding@home, World Community Grid. Координує мільйони ПК добровольців для наукових обчислень. Низька надійність вузлів компенсується перевіркою результатів (quorum).'));

content.push(h2('49.4. HTCondor'));
content.push(p('HTCondor (High Throughput Computing) — система управління задачами для кластерів та Grid. ClassAd механізм для матчингу задач та ресурсів. DAGMan — планування задач з залежностями.'));

content.push(h2('49.5. UNICORE'));
content.push(p('UNICORE (Uniform Interface to Computing Resources) — Grid middleware, розроблений в Європі. Використовується в DEISA та PRACE (Partnership for Advanced Computing in Europe). Надає уніфікований доступ до суперкомп\'ютерів різних центрів.'));

content.push(h2('49.6. Зв\'язок з Java'));
content.push(p('Java широко використовується в Grid ПЗ: GRAM клієнти, GridFTP API, Web Services (JAX-WS). Spring Boot мікросервіси можуть бути компонентами Grid-застосунків. Java\'s platform independence є перевагою в гетерогенних Grid середовищах.'));


// ─── TITLE PAGE ─────────────────────────────────────────────────────────────

const titlePage = [
  new Paragraph({
    children: [new TextRun({ text: '', size: 24 })],
    spacing: { before: 2880, after: 0 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'ТЕХНОЛОГІЯ ПАРАЛЕЛЬНИХ ОБЧИСЛЕНЬ', bold: true, size: 40, font: 'Arial' })],
    spacing: { before: 0, after: 240 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Відповіді на екзаменаційні питання', size: 32, font: 'Arial', color: '444444' })],
    spacing: { before: 0, after: 240 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Мова реалізації: Java', size: 28, font: 'Arial', color: '666666', italics: true })],
    spacing: { before: 0, after: 480 },
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: 'Питання: 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47, 49', size: 22, font: 'Arial', color: '888888' })
    ],
    spacing: { before: 0, after: 0 },
  }),
];

// ─── BUILD DOCUMENT ──────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 24 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: '2E75B6' },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      },
      {
        id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: '2E4057' },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: 'numbers',
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
    ]
  },
  sections: [
    // Title page
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
        }
      },
      children: titlePage,
    },
    // Main content
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Технологія паралельних обчислень — Відповіді на екзаменаційні питання', size: 18, font: 'Arial', color: '888888' })
              ],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
              spacing: { after: 200 },
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({
                  text: 'Технологія паралельних обчислень',
                  size: 18,
                  font: 'Arial',
                  color: '888888',
                }),
              ],
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
            })
          ]
        })
      },
      children: content,
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('./docs/uv4.docx', buffer);
}).catch(err => {
  console.error('❌ Помилка:', err);
});