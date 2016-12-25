<?php
namespace app\index\controller;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\ListItem;
use PhpOffice\PhpWord\Style\TOC;

class Index
{
    public function index()
    {
//        return view('admin@main/main');
        return view('index/index');
    }

    /**
     * 生成Word文件
     *
     */
    public function generateWord()
    {
        $phpWord = new PhpWord();

        $section = $phpWord->addSection();

        //预定义样式
        $phpWord->setDefaultFontName("宋体");         //设置默认字体宋体
        $phpWord->setDefaultFontSize(12);               //小四
        $styleTOC = array('tabLeader' => TOC::TAB_LEADER_DOT);
        $phpWord->addTitleStyle(2, array('size' => 16, 'bold' => true, 'name' => '宋体'));

        //添加页脚
        $footer = $section->addFooter();
        $footer->addPreserveText("{PAGE}", array(), array('alignment' => Jc::CENTER));

        //目录
        $section->addText("目录", array(), array('alignment' => Jc::CENTER));
        $section->addTOC(array(), $styleTOC);

        $section->addPageBreak();

        $section->addTitle("一、毕业实习的课题背景", 2);
        $text1 = "随着科学技术和移动互联网的发展，二十一世纪已经迈入了一个集数字化，网络化，信息化的时代。";
        $section->addText($text1,
            null,
            array('indentation' => array('firstLine' => 480))
        );
        $text2 = "面对激烈的竞争环境，餐饮行业的管理需要更加的规范化和科学化。";
        $section->addText($text2,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle("二、毕业实习的技术参数（研究内容）", 2);
        $text3 = "本次实习任务是研究实现一套完整的手机饭店点菜系统，包括网站服务端和移动Android端。";
        $section->addText($text3,
            null,
            array('indentation' => array('firstLine' => 480))
        );
        $section->addText("具体分为以下步骤：",
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle("三、毕业实习应完成的具体工作及提交形式", 2);
        $text4_1 = "1. 查找相关资料，对自己的选题做初步的了解；";
        $text4_2 = "2. 制定初步计划，完成任务书；";
        $text4_3 = "3. 收集资料并进行筛选；";
        $section->addText($text4_1,
            null,
            array('indentation' => array('firstLine' => 480))
        );
        $section->addText($text4_2,
            null,
            array('indentation' => array('firstLine' => 480))
        );
        $section->addText($text4_3,
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle("四、毕业实习进度安排", 2);
        $section->addText("这里没有任何的文字哦！",
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $section->addPageBreak();

        $section->addTitle("五、应收集的资料及主要参考文献", 2);
        $section->addText("[1] 郭宏志. Android应用开发详解[M]. 电子工业出版社, 2010.",
            null,
            array('indentation' => array('firstLine' => 480))
        );

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }

    public function testWord()
    {
        $phpWord = new PhpWord();

        //预定义样式
        $phpWord->setDefaultFontName("宋体");         //设置默认字体宋体
        $phpWord->setDefaultFontSize(12);               //小四
        $levelStyle = array('listType' => ListItem::TYPE_NUMBER_NESTED);
        $pStyle = array('indentation' => array('firstLine' => 480));

        // New section
        $section = $phpWord->addSection();

        $section->addText("多层次的缩进级别。", array(), $pStyle);
        $section->addListItem("天下武功，唯快不破", 0, array(), $levelStyle,
            array('indentation' => array('firstLine' => 480))
        );
        $section->addListItem("天下武功，唯快不破", 0, array(), $levelStyle,
            array('indentation' => array('firstLine' => 480))
        );
        $section->addTextBreak(2);

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }

    public function testWord2()
    {
        $phpWord = new PhpWord();

        //预定义样式
        $phpWord->setDefaultFontName("宋体");         //设置默认字体宋体
        $phpWord->setDefaultFontSize(12);               //小四

        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $phpWord->setDefaultParagraphStyle(
            array(
                'alignment' => \PhpOffice\PhpWord\SimpleType\Jc::BOTH,
                'spaceAfter' => \PhpOffice\PhpWord\Shared\Converter::pointToTwip(12),
                'spacing' => 120,
            )
        );

// New section
        $section = $phpWord->addSection();

        $section->addText(
            'Below are the samples on how to control your paragraph '
            . 'pagination. See "Line and Page Break" tab on paragraph properties '
            . 'window to see the attribute set by these controls.',
            array('bold' => true),
            array('space' => array('before' => 360, 'after' => 480))
        );

        $section->addText(
            'Paragraph with widowControl = false (default: true). '
            . 'A "widow" is the last line of a paragraph printed by itself at the top '
            . 'of a page. An "orphan" is the first line of a paragraph printed by '
            . 'itself at the bottom of a page. Set this option to "false" if you want '
            . 'to disable this automatic control.',
            null,
            array('widowControl' => false, 'indentation' => array('left' => 240, 'right' => 120))
        );

        $section->addText(
            'Paragraph with keepNext = true (default: false). '
            . '"Keep with next" is used to prevent Word from inserting automatic page '
            . 'breaks between paragraphs. Set this option to "true" if you do not want '
            . 'your paragraph to be separated with the next paragraph.',
            null,
            array('keepNext' => true, 'indentation' => array('firstLine' => 240))
        );

        $section->addText(
            'Paragraph with keepLines = true (default: false). '
            . '"Keep lines together" will prevent Word from inserting an automatic page '
            . 'break within a paragraph. Set this option to "true" if you do not want '
            . 'all lines of your paragraph to be in the same page.',
            null,
            array('keepLines' => true, 'indentation' => array('left' => 240, 'hanging' => 240))
        );

        $section->addText('Keep scrolling. More below.');

        $section->addText(
            '一个点菜网站系统包括顾客点菜，厨师炒菜，老板结账，以及系统管理等几部分，每部分都通过智能手机上网操作和流转，整个业务融为一体。',
            array(),
            array('indentation' => array('firstLine' => 480))
        );

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }

    public function testWord3()
    {
        $phpWord = new PhpWord();

        // Define styles
        $fontStyleName = 'myOwnStyle';
        $phpWord->addFontStyle($fontStyleName, array('color' => 'FF0000'));

        $paragraphStyleName = 'P-Style';
        $phpWord->addParagraphStyle($paragraphStyleName, array('spaceAfter' => 95));

        $multilevelNumberingStyleName = 'multilevel';
        $phpWord->addNumberingStyle(
            $multilevelNumberingStyleName,
            array(
                'type' => 'multilevel',
                'levels' => array(
                    array('format' => 'decimal', 'text' => '%1.', 'left' => 360, 'hanging' => 360, 'tabPos' => 360),
                    array('format' => 'upperLetter', 'text' => '%2.', 'left' => 720, 'hanging' => 360, 'tabPos' => 720),
                ),
            )
        );

        $predefinedMultilevelStyle = array('listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER_NESTED);

// New section
        $section = $phpWord->addSection();

// Lists
        $section->addText('Multilevel list.');
        $section->addListItem('List Item I', 0, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item I.a', 1, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item I.b', 1, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item II', 0, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item II.a', 1, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item III', 0, null, $multilevelNumberingStyleName);
        $section->addTextBreak(2);

        $section->addText('Basic simple bulleted list.');
        $section->addListItem('List Item 1');
        $section->addListItem('List Item 2');
        $section->addListItem('List Item 3');
        $section->addTextBreak(2);

        $section->addText('Continue from multilevel list above.');
        $section->addListItem('List Item IV', 0, null, $multilevelNumberingStyleName);
        $section->addListItem('List Item IV.a', 1, null, $multilevelNumberingStyleName);
        $section->addTextBreak(2);

        $section->addText('Multilevel predefined list.');
        $section->addListItem('List Item 1', 0, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 2', 0, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 3', 1, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 4', 1, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 5', 2, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 6', 1, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addListItem('List Item 7', 0, $fontStyleName, $predefinedMultilevelStyle, $paragraphStyleName);
        $section->addTextBreak(2);

        $section->addText('List with inline formatting.');
        $listItemRun = $section->addListItemRun();
        $listItemRun->addText('List item 1');
        $listItemRun->addText(' in bold', array('bold' => true));
        $listItemRun = $section->addListItemRun();
        $listItemRun->addText('List item 2');
        $listItemRun->addText(' in italic', array('italic' => true));
        $listItemRun = $section->addListItemRun();
        $listItemRun->addText('List item 3');
        $listItemRun->addText(' underlined', array('underline' => 'dash'));
        $section->addTextBreak(2);

// Numbered heading
        $headingNumberingStyleName = 'headingNumbering';
        $phpWord->addNumberingStyle(
            $headingNumberingStyleName,
            array(
                'type' => 'multilevel',
                'levels' => array(
                    array('pStyle' => 'Heading1', 'format' => 'decimal', 'text' => '%1'),
                    array('pStyle' => 'Heading2', 'format' => 'decimal', 'text' => '%1.%2'),
                    array('pStyle' => 'Heading3', 'format' => 'decimal', 'text' => '%1.%2.%3'),
                ),
            )
        );
        $phpWord->addTitleStyle(1, array('size' => 16),
            array('numStyle' => $headingNumberingStyleName, 'numLevel' => 0));
        $phpWord->addTitleStyle(2, array('size' => 14),
            array('numStyle' => $headingNumberingStyleName, 'numLevel' => 1));
        $phpWord->addTitleStyle(3, array('size' => 12),
            array('numStyle' => $headingNumberingStyleName, 'numLevel' => 2));

        $section->addTitle('Heading 1', 1);
        $section->addTitle('Heading 2', 2);
        $section->addTitle('Heading 3', 3);

        $fileName = "论文小助手Pro" . date("YmdHis") . ".docx";
        $phpWord->save($fileName, "Word2007", true);
    }
}
