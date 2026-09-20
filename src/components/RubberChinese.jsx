import { useEffect, useMemo, useState } from 'react';
import { Search, Volume2 } from 'lucide-react';
import rubberChineseData from '../data/rubber-chinese.json';

const ITEMS_PER_PAGE = 9;

function speakChinese(text) {
    if (typeof window === 'undefined') return;
    if (!window.speechSynthesis) return;

    window.speechSynthesis.cancel();

    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.8;
    utterance.pitch = 1;

    window.speechSynthesis.speak(utterance);
}

export default function RubberChinese() {
    const [categoryId, setCategoryId] = useState('all');
    const [keyword, setKeyword] = useState('');
    const [currentPage, setCurrentPage] = useState(1);

    const categories = Array.isArray(rubberChineseData?.categories)
        ? rubberChineseData.categories
        : [];

    /*
     * ============================
     * LỌC DỮ LIỆU
     * ============================
     */
    const words = useMemo(() => {
        let result = [];

        // Tất cả category
        if (categoryId === 'all') {
            result = categories.flatMap(category =>
                Array.isArray(category.words)
                    ? category.words
                    : []
            );
        }

        // Category được chọn
        else {
            const category = categories.find(
                item => item.id === categoryId
            );

            result = Array.isArray(category?.words)
                ? category.words
                : [];
        }

        // Tìm kiếm
        const searchText = keyword.trim().toLowerCase();

        if (!searchText) {
            return result;
        }

        return result.filter(word => {
            const values = [
                word.hanzi,
                word.pinyin,
                word.vietnamese,
                word.hanviet
            ];

            return values.some(value =>
                String(value || '')
                    .toLowerCase()
                    .includes(searchText)
            );
        });
    }, [categories, categoryId, keyword]);

    /*
     * ============================
     * PHÂN TRANG
     * ============================
     */
    const totalPages = Math.ceil(
        words.length / ITEMS_PER_PAGE
    );

    const paginatedWords = useMemo(() => {
        const startIndex =
            (currentPage - 1) * ITEMS_PER_PAGE;

        return words.slice(
            startIndex,
            startIndex + ITEMS_PER_PAGE
        );
    }, [words, currentPage]);

    /*
     * Khi tìm kiếm hoặc đổi category
     * → quay về trang 1
     */
    useEffect(() => {
        setCurrentPage(1);
    }, [categoryId, keyword]);

    /*
     * Nếu số trang giảm xuống
     * → đảm bảo currentPage không vượt quá totalPages
     */
    useEffect(() => {
        if (totalPages > 0 && currentPage > totalPages) {
            setCurrentPage(totalPages);
        }
    }, [currentPage, totalPages]);

    /*
     * ============================
     * ĐỔI TRANG
     * ============================
     */
    const goToPage = page => {
        if (page < 1 || page > totalPages) {
            return;
        }

        setCurrentPage(page);

        // Cuộn lên đầu danh sách
        window.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
    };

    return (
        <div className="w-full">

            {/* =========================================
                HEADER
            ========================================= */}
            <header className="mb-6 sm:mb-7">

                <h1 className="text-xl sm:text-[22px] font-bold tracking-tight">
                    Tiếng Trung luyện su
                </h1>

                <p className="text-ink-soft text-sm sm:text-base mt-1.5">
                    Từ vựng tiếng Trung chuyên ngành luyện su
                </p>

            </header>


            {/* =========================================
                SEARCH
            ========================================= */}
            <div className="relative mb-5">

                <Search
                    size={16}
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-ink-faint"
                />

                <input
                    type="text"
                    value={keyword}
                    onChange={event =>
                        setKeyword(event.target.value)
                    }
                    placeholder="Tìm kiếm từ vựng..."
                    className="
                        w-full
                        bg-surface
                        border
                        border-line
                        rounded-lg
                        pl-9
                        pr-4
                        py-2.5
                        text-sm
                        outline-none
                        focus:border-accent
                        transition-colors
                    "
                />

            </div>


            {/* =========================================
                CATEGORY
            ========================================= */}
            <div className="
                flex
                flex-wrap
                gap-2
                mb-5
            ">

                {/* Tất cả */}
                <button
                    type="button"
                    onClick={() => {
                        setCategoryId('all');
                        setCurrentPage(1);
                    }}
                    className={`
                        px-3
                        py-2
                        rounded-lg
                        text-sm
                        font-medium
                        transition-colors
                        whitespace-nowrap

                        ${
                            categoryId === 'all'
                                ? 'bg-accent text-[#241605]'
                                : 'bg-surface border border-line text-ink-soft hover:border-accent'
                        }
                    `}
                >
                    Tất cả
                </button>


                {/* Category */}
                {categories.map(category => (

                    <button
                        type="button"
                        key={category.id}
                        onClick={() => {
                            setCategoryId(category.id);
                            setCurrentPage(1);
                        }}
                        className={`
                            px-3
                            py-2
                            rounded-lg
                            text-sm
                            font-medium
                            transition-colors
                            whitespace-nowrap

                            ${
                                categoryId === category.id
                                    ? 'bg-accent text-[#241605]'
                                    : 'bg-surface border border-line text-ink-soft hover:border-accent'
                            }
                        `}
                    >
                        {category.icon} {category.nameVi}
                    </button>

                ))}

            </div>


            {/* =========================================
                RESULT COUNT
            ========================================= */}
            <div className="
                flex
                flex-col
                sm:flex-row
                sm:items-center
                sm:justify-between
                gap-1
                mb-4
                text-xs
                text-ink-faint
            ">

                <span>
                    {words.length} từ vựng
                </span>

                {words.length > 0 && (
                    <span>
                        Trang {currentPage} / {totalPages}
                    </span>
                )}

            </div>


            {/* =========================================
                EMPTY
            ========================================= */}
            {words.length === 0 ? (

                <div className="
                    bg-surface
                    border
                    border-line
                    rounded-xl
                    p-8
                    text-center
                ">
                    <p className="text-ink-soft">
                        Không tìm thấy từ vựng phù hợp.
                    </p>
                </div>

            ) : (

                <>

                    {/* =================================
                        WORD CARDS
                    ================================= */}
                    <div className="
                        grid
                        grid-cols-1
                        sm:grid-cols-2
                        lg:grid-cols-3
                        gap-4
                    ">

                        {paginatedWords.map((word, index) => (

                            <div
                                key={`${word.hanzi}-${index}`}
                                className="
                                    bg-surface
                                    border
                                    border-line
                                    rounded-xl
                                    p-4
                                    sm:p-5
                                    shadow-card
                                    transition-all
                                    hover:-translate-y-0.5
                                "
                            >

                                {/* Hanzi + Audio */}
                                <div className="
                                    flex
                                    items-start
                                    justify-between
                                    gap-3
                                ">

                                    <div className="min-w-0">

                                        <div className="
                                            text-[28px]
                                            sm:text-[30px]
                                            font-bold
                                            leading-tight
                                            break-words
                                        ">
                                            {word.hanzi}
                                        </div>

                                        <div className="
                                            text-accent-dark
                                            font-medium
                                            mt-2
                                            text-sm
                                            sm:text-base
                                        ">
                                            {word.pinyin}
                                        </div>

                                    </div>


                                    {/* Audio */}
                                    <button
                                        type="button"
                                        onClick={() =>
                                            speakChinese(word.hanzi)
                                        }
                                        title="Nghe phát âm"
                                        aria-label={`Nghe phát âm ${word.hanzi}`}
                                        className="
                                            w-9
                                            h-9
                                            shrink-0
                                            rounded-lg
                                            bg-accent-soft
                                            text-accent-dark
                                            flex
                                            items-center
                                            justify-center
                                            hover:bg-accent
                                            transition-colors
                                        "
                                    >
                                        <Volume2 size={17} />
                                    </button>

                                </div>


                                {/* Vietnamese */}
                                <div className="
                                    mt-4
                                    pt-3
                                    border-t
                                    border-line
                                ">

                                    <div className="
                                        text-sm
                                        font-semibold
                                        break-words
                                    ">
                                        {word.vietnamese}
                                    </div>


                                    {word.hanviet && (

                                        <div className="
                                            text-xs
                                            text-ink-faint
                                            mt-1
                                            break-words
                                        ">
                                            Hán-Việt: {word.hanviet}
                                        </div>

                                    )}

                                </div>

                            </div>

                        ))}

                    </div>


                    {/* =================================
                        PAGINATION
                    ================================= */}
                    {totalPages > 1 && (

                        <div className="
                            flex
                            items-center
                            justify-center
                            gap-1.5
                            mt-7
                            flex-wrap
                        ">

                            {/* Previous */}
                            <button
                                type="button"
                                disabled={currentPage === 1}
                                onClick={() =>
                                    goToPage(currentPage - 1)
                                }
                                className="
                                    min-w-9
                                    h-9
                                    px-2
                                    rounded-lg
                                    border
                                    border-line
                                    bg-surface
                                    text-sm
                                    disabled:opacity-40
                                    disabled:cursor-not-allowed
                                    hover:border-accent
                                    transition-colors
                                "
                                aria-label="Trang trước"
                            >
                                ←
                            </button>


                            {/* Page numbers */}
                            {Array.from(
                                { length: totalPages },
                                (_, index) => {

                                    const page = index + 1;

                                    return (

                                        <button
                                            key={page}
                                            type="button"
                                            onClick={() =>
                                                goToPage(page)
                                            }
                                            className={`
                                                min-w-9
                                                h-9
                                                px-2
                                                rounded-lg
                                                text-sm
                                                font-medium
                                                transition-colors

                                                ${
                                                    currentPage === page
                                                        ? 'bg-accent text-[#241605]'
                                                        : 'bg-surface border border-line text-ink-soft hover:border-accent'
                                                }
                                            `}
                                        >
                                            {page}
                                        </button>

                                    );
                                }
                            )}


                            {/* Next */}
                            <button
                                type="button"
                                disabled={
                                    currentPage === totalPages
                                }
                                onClick={() =>
                                    goToPage(currentPage + 1)
                                }
                                className="
                                    min-w-9
                                    h-9
                                    px-2
                                    rounded-lg
                                    border
                                    border-line
                                    bg-surface
                                    text-sm
                                    disabled:opacity-40
                                    disabled:cursor-not-allowed
                                    hover:border-accent
                                    transition-colors
                                "
                                aria-label="Trang sau"
                            >
                                →
                            </button>

                        </div>

                    )}

                </>

            )}

        </div>
    );
}